/**
 * Reprices historical LibreChat transactions after a price-table correction.
 *
 * WHY THIS EXISTS
 * LibreChat bills each message using its own bundled price table and stores the
 * result on the transaction. When a model is missing from that table it is
 * priced at a catch-all rate instead of failing, so the spend figures in
 * Organization Settings are quietly wrong — claude-sonnet-5 was billed at
 * $0.80/$2.40 instead of $3/$15 for two weeks before anyone noticed. Fixing the
 * table (see patch.js) only affects messages sent afterwards; everything
 * already recorded keeps the wrong value. This script restates that history.
 *
 * It is also the procedure to run when Anthropic changes a price: update RATES,
 * dry-run, then apply over the affected date range.
 *
 * SAFETY
 *   - Dry run by default. Nothing is written without --apply.
 *   - The original values are preserved on each row (tokenValueOriginal,
 *     rateOriginal, repricedAt), so a run can be inspected or undone.
 *   - Idempotent: rows already repriced are skipped, so re-running cannot
 *     compound values.
 *
 * WHERE TO RUN IT
 * Inside the bridge container. The Mongo Atlas cluster is not reachable from a
 * developer machine (its SRV lookup is blocked), and the container already has
 * MONGO_URI set, so this lives under src/ to be compiled into dist/ rather than
 * in scripts/, which the build excludes:
 *
 *   az containerapp exec -n agent365-bridge -g ABKAgent365 --command \
 *     "node /app/dist/maintenance/reprice-transactions.js --from 2026-09-01 --to 2026-09-02"
 *
 * Add --apply to commit, and --model <id> to limit it to one model.
 */
import { MongoClient, Document } from "mongodb";

/**
 * Anthropic list prices, USD per 1M tokens. Keep in sync with the table
 * injected by patch.js — this script and that patch must agree, or repriced
 * history will disagree with newly recorded messages.
 */
const RATES: Record<string, { prompt: number; completion: number }> = {
  "claude-fable-5":   { prompt: 10, completion: 50 },
  "claude-opus-5":    { prompt: 5,  completion: 25 },
  "claude-opus-4-8":  { prompt: 5,  completion: 25 },
  "claude-opus-4-7":  { prompt: 5,  completion: 25 },
  "claude-opus-4-6":  { prompt: 5,  completion: 25 },
  "claude-opus-4-5":  { prompt: 5,  completion: 25 },
  "claude-sonnet-5":  { prompt: 3,  completion: 15 },
  "claude-sonnet-4-6": { prompt: 3, completion: 15 },
  "claude-sonnet-4-5": { prompt: 3, completion: 15 },
  "claude-haiku-4-5": { prompt: 1,  completion: 5 },
};

/** Dated model ids (claude-sonnet-4-5-20250929) price as their base model. */
function ratesFor(model: string): { prompt: number; completion: number } | null {
  if (RATES[model]) return RATES[model];
  const key = Object.keys(RATES)
    .filter((k) => model.startsWith(k))
    .sort((a, b) => b.length - a.length)[0];
  return key ? RATES[key] : null;
}

function arg(name: string): string | undefined {
  const i = process.argv.indexOf(`--${name}`);
  return i !== -1 ? process.argv[i + 1] : undefined;
}

function parseDate(v: string | undefined, label: string): Date {
  if (!v) throw new Error(`--${label} is required (YYYY-MM-DD)`);
  const d = new Date(`${v}T00:00:00.000Z`);
  if (Number.isNaN(d.getTime())) throw new Error(`--${label} is not a valid date: ${v}`);
  return d;
}

async function main() {
  const uri = process.env.MONGO_URI;
  if (!uri) throw new Error("MONGO_URI is not set.");

  const from = parseDate(arg("from"), "from");
  const to = parseDate(arg("to"), "to");
  const onlyModel = arg("model");
  const apply = process.argv.includes("--apply");

  const client = new MongoClient(uri);
  await client.connect();
  const col = client.db().collection("transactions");

  const query: Document = {
    createdAt: { $gte: from, $lt: to },
    repricedAt: { $exists: false },
  };
  if (onlyModel) query.model = onlyModel;

  const rows = await col.find(query).toArray();
  console.log(`Range ${from.toISOString().slice(0, 10)} -> ${to.toISOString().slice(0, 10)}`);
  console.log(`${rows.length} transaction(s) not yet repriced${onlyModel ? ` for ${onlyModel}` : ""}\n`);

  let oldTotal = 0;
  let newTotal = 0;
  let changed = 0;
  const unknown = new Set<string>();
  const perModel: Record<string, { old: number; neu: number; rows: number }> = {};

  for (const r of rows) {
    const oldUsd = Math.abs((r.tokenValue as number) ?? 0) / 1e6;
    oldTotal += oldUsd;

    const model = String(r.model ?? "");
    const rate = ratesFor(model);
    if (!rate) {
      // Never guess: an unpriced model is reported so it can be added to both
      // RATES and patch.js, rather than silently left at whatever it holds.
      unknown.add(model);
      newTotal += oldUsd;
      continue;
    }

    const perMillion = r.tokenType === "completion" ? rate.completion : rate.prompt;
    const tokens = Math.abs((r.rawAmount as number) ?? 0);
    const newValue = tokens * perMillion;
    const newUsd = newValue / 1e6;
    newTotal += newUsd;

    const m = (perModel[model] ??= { old: 0, neu: 0, rows: 0 });
    m.old += oldUsd;
    m.neu += newUsd;
    m.rows++;

    if (apply) {
      await col.updateOne(
        { _id: r._id },
        {
          $set: {
            tokenValue: -newValue,
            rate: perMillion,
            tokenValueOriginal: r.tokenValue,
            rateOriginal: r.rate,
            repricedAt: new Date(),
          },
        }
      );
    }
    changed++;
  }

  for (const [model, v] of Object.entries(perModel)) {
    console.log(`  ${model.padEnd(30)} ${String(v.rows).padStart(5)} rows   $${v.old.toFixed(4)} -> $${v.neu.toFixed(4)}`);
  }
  if (unknown.size) {
    console.log(`\n  NOT PRICED (left untouched): ${[...unknown].join(", ")}`);
    console.log("  Add these to RATES here and to patch.js, then re-run.");
  }

  const factor = oldTotal > 0 ? (newTotal / oldTotal).toFixed(2) : "n/a";
  console.log(`\nTOTAL  $${oldTotal.toFixed(4)} -> $${newTotal.toFixed(4)}  (x${factor})`);
  console.log(apply
    ? `Applied to ${changed} transaction(s). Originals kept in tokenValueOriginal / rateOriginal.`
    // Output stays strictly ASCII: this runs through `az containerapp exec`,
    // whose reader encodes to the terminal's codepage. On a Greek Windows
    // console (cp1253) a single "->" arrow or em dash raises UnicodeEncodeError
    // and kills the session mid-run, hiding the result entirely.
    : `Dry run - nothing written. Re-run with --apply to commit.`);

  await client.close();
}

main().catch((e) => {
  console.error("Failed:", e instanceof Error ? e.message : e);
  process.exit(1);
});
