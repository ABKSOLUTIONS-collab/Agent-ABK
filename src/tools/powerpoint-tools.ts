/**
 * powerpoint-tools.ts
 *
 * Reads PowerPoint decks (.pptx) from OneDrive or SharePoint.
 *
 * A .pptx is a ZIP of XML parts, so the text can be pulled out directly with
 * fflate — the same approach word-graph-tools already uses for .docx. No
 * Copilot licence and no extra Graph permission are involved: the file is
 * fetched with the delegated token the user already granted.
 */
import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { strFromU8, unzipSync } from "fflate";
import { fetchDriveItem } from "./sharepoint-tools";

function log(msg: string): void {
  process.stderr.write(`[agent365-bridge] [powerpoint] ${msg}\n`);
}

export const POWERPOINT_TOOL_NAMES = new Set(["GetPresentationContent"]);

export const POWERPOINT_TOOLS: Tool[] = [
  {
    name: "GetPresentationContent",
    description:
      "Read the text of a PowerPoint presentation (.pptx) from the user's OneDrive or a SharePoint " +
      "site, slide by slide. Use this to summarise a deck, answer questions about it, or pull out " +
      "specific slides. Identify the file by path (e.g. '/Decks/Q3 Review.pptx') or by the item ID " +
      "returned from a folder listing. Speaker notes are included by default.",
    inputSchema: {
      type: "object",
      properties: {
        file: {
          type: "string",
          description: "The presentation: a path relative to the drive root, or its item ID.",
        },
        siteId: {
          type: "string",
          description: "Optional. SharePoint site ID, when the deck lives in a site rather than the user's OneDrive.",
        },
        includeNotes: {
          type: "boolean",
          description: "Include speaker notes for each slide. Default: true.",
        },
      },
      required: ["file"],
    },
  },
];

// Slide parts are named slide1.xml, slide2.xml, ... slide10.xml — plain string
// ordering would put slide10 before slide2, so sort on the trailing number.
function slideNumber(name: string): number {
  const m = name.match(/(\d+)\.xml$/);
  return m ? Number(m[1]) : 0;
}

// Drawing-ML holds every run of visible text in <a:t> elements, whatever the
// shape it belongs to (title, body, table cell, chart label). Collecting those
// gets the slide's text without having to model the shape tree.
function extractPartText(xml: string): string {
  const runs: string[] = [];
  for (const m of xml.matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)) {
    const t = m[1].trim();
    if (t) runs.push(t);
  }
  return runs.join(" ").replace(/\s+/g, " ").trim();
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

export class PowerPointToolHandler {
  private token: string;

  constructor(token: string) {
    this.token = token;
  }

  async handleToolCall(
    toolName: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
    try {
      if (toolName !== "GetPresentationContent") {
        return { content: [{ type: "text", text: `Unknown tool: ${toolName}` }], isError: true };
      }

      const ref = String(args.file ?? "").trim();
      if (!ref) {
        return { content: [{ type: "text", text: "Error: 'file' is required." }], isError: true };
      }
      const includeNotes = args.includeNotes !== false;

      const item = await fetchDriveItem(this.token, ref, args.siteId as string | undefined);
      if (!/\.pptx$/i.test(item.name)) {
        return {
          content: [{ type: "text", text: `'${item.name}' is not a .pptx file. Only PowerPoint presentations in the modern format can be read; a legacy .ppt must be converted first.` }],
          isError: true,
        };
      }

      const files = unzipSync(item.buffer);
      const slideNames = Object.keys(files)
        .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
        .sort((a, b) => slideNumber(a) - slideNumber(b));

      if (slideNames.length === 0) {
        return { content: [{ type: "text", text: `'${item.name}' contains no slides.` }], isError: true };
      }

      const sections: string[] = [];
      for (const name of slideNames) {
        const n = slideNumber(name);
        const body = decodeXmlEntities(extractPartText(strFromU8(files[name])));

        let notes = "";
        if (includeNotes) {
          // Notes live in a parallel part sharing the slide's number.
          const notesPart = files[`ppt/notesSlides/notesSlide${n}.xml`];
          if (notesPart) {
            const raw = decodeXmlEntities(extractPartText(strFromU8(notesPart)));
            // The notes part repeats the slide number as a standalone run;
            // dropping an exact numeric match avoids a spurious "Notes: 3".
            if (raw && raw !== String(n)) notes = raw;
          }
        }

        sections.push(
          `--- Slide ${n} ---\n${body || "(no text on this slide)"}${notes ? `\n[Speaker notes] ${notes}` : ""}`
        );
      }

      log(`Read '${item.name}': ${slideNames.length} slides`);
      return {
        content: [{ type: "text", text: `${item.name} — ${slideNames.length} slide(s)\n\n${sections.join("\n\n")}` }],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      log(`Error (${toolName}): ${message}`);
      return { content: [{ type: "text", text: `Error: ${message}` }], isError: true };
    }
  }
}
