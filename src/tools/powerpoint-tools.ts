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
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { fetchDriveItem, buildUploadPath } from "./sharepoint-tools";

function log(msg: string): void {
  process.stderr.write(`[agent365-bridge] [powerpoint] ${msg}\n`);
}

export const POWERPOINT_TOOL_NAMES = new Set([
  "GetPresentationContent",
  "CreatePresentation",
  "EditPresentation",
]);

// ── Reading an existing package ───────────────────────────────────────────────

type Parts = Record<string, Uint8Array>;

/**
 * The slides of a deck in the order PowerPoint displays them.
 *
 * That order lives in <p:sldIdLst> inside presentation.xml, NOT in the part
 * file names — a deck that has had slides reordered or deleted routinely has
 * slide3.xml appearing second. Anything that edits by slide number has to
 * resolve through this list, or it will silently modify the wrong slide.
 */
function orderedSlides(files: Parts): Array<{ sldId: string; rId: string; part: string }> {
  const pres = strFromU8(files["ppt/presentation.xml"] ?? new Uint8Array());
  const rels = strFromU8(files["ppt/_rels/presentation.xml.rels"] ?? new Uint8Array());

  const targets = new Map<string, string>();
  for (const m of rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g)) {
    targets.set(m[1], m[2]);
  }

  const out: Array<{ sldId: string; rId: string; part: string }> = [];
  const listMatch = pres.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
  if (!listMatch) return out;
  for (const m of listMatch[1].matchAll(/<p:sldId\b[^>]*id="(\d+)"[^>]*r:id="([^"]+)"[^>]*\/>/g)) {
    const target = targets.get(m[2]);
    if (!target) continue;
    out.push({ sldId: m[1], rId: m[2], part: `ppt/${target.replace(/^\.\//, "")}` });
  }
  return out;
}

function nextNumericId(values: number[], floor: number): number {
  return Math.max(floor - 1, ...values) + 1;
}

// ── .pptx package builder ─────────────────────────────────────────────────────
//
// A presentation has no Graph API equivalent to Excel's workbook endpoints, so
// a new deck has to be assembled as a complete OOXML package. PowerPoint will
// refuse to open a file that is missing any of the parts below — the master,
// layout and theme are mandatory even for a one-slide deck, which is why this
// is more boilerplate than the equivalent Word or Excel builders.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export interface SlideSpec {
  title?: string;
  bullets?: string[];
  notes?: string;
}

function slideXml(spec: SlideSpec): string {
  const title = esc(spec.title ?? "");
  const bullets = (spec.bullets ?? []).filter((b) => String(b).trim());

  const bulletParas = bullets.length
    ? bullets.map((b) => `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(String(b))}</a:t></a:r></a:p>`).join("")
    : `<a:p><a:endParaRPr lang="en-US"/></a:p>`;

  // EMU units: 914400 per inch. These place a title band and a body block on a
  // standard 13.33 x 7.5 inch (16:9) slide.
  return `${XML_DECL}
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${title}</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/>${bulletParas}</p:txBody>
</p:sp>
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function notesXml(notes: string): string {
  return `${XML_DECL}
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(notes)}</a:t></a:r></a:p></p:txBody>
</p:sp>
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

function buildPptx(slides: SlideSpec[], title: string): Buffer {
  const n = slides.length;
  const slideIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  const slideOverrides = slides.map((_, i) =>
    `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  ).join("");
  const notesOverrides = slides.map((s, i) => s.notes
    ? `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`
    : ""
  ).join("");

  const files: Record<string, Uint8Array> = {};

  files["[Content_Types].xml"] = strToU8(`${XML_DECL}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
${slideOverrides}${notesOverrides}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);

  files["_rels/.rels"] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);

  // 12192000 x 6858000 EMU = 13.33 x 7.5 inches, the 16:9 default.
  files["ppt/presentation.xml"] = strToU8(`${XML_DECL}
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>${slideIds}</p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`);

  const presRels = slides.map((_, i) =>
    `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
  ).join("");
  files["ppt/_rels/presentation.xml.rels"] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
${presRels}
<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);

  const placeholderTree = (bodyIdx: string) => `<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${bodyIdx}/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
</p:spTree>`;

  files["ppt/slideMasters/slideMaster1.xml"] = strToU8(`${XML_DECL}
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>${placeholderTree('type="body" idx="1"')}</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`);

  files["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);

  files["ppt/slideLayouts/slideLayout1.xml"] = strToU8(`${XML_DECL}
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj" preserve="1">
<p:cSld name="Title and Content">${placeholderTree('idx="1"')}</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`);

  files["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

  // PowerPoint requires a theme with all three schemes present; these are the
  // Office defaults, trimmed to the minimum it will accept.
  const dk = (n2: string, v: string) => `<a:${n2}><a:srgbClr val="${v}"/></a:${n2}>`;
  files["ppt/theme/theme1.xml"] = strToU8(`${XML_DECL}
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
<a:themeElements>
<a:clrScheme name="Office">${dk("dk1", "000000")}${dk("lt1", "FFFFFF")}${dk("dk2", "44546A")}${dk("lt2", "E7E6E6")}${dk("accent1", "4472C4")}${dk("accent2", "ED7D31")}${dk("accent3", "A5A5A5")}${dk("accent4", "FFC000")}${dk("accent5", "5B9BD5")}${dk("accent6", "70AD47")}${dk("hlink", "0563C1")}${dk("folHlink", "954F72")}</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`);

  slides.forEach((spec, i) => {
    const idx = i + 1;
    files[`ppt/slides/slide${idx}.xml`] = strToU8(slideXml(spec));
    const rels = [`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`];
    if (spec.notes) {
      rels.push(`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${idx}.xml"/>`);
      files[`ppt/notesSlides/notesSlide${idx}.xml`] = strToU8(notesXml(spec.notes));
      files[`ppt/notesSlides/_rels/notesSlide${idx}.xml.rels`] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${idx}.xml"/>
</Relationships>`);
    }
    files[`ppt/slides/_rels/slide${idx}.xml.rels`] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`);
  });

  files["docProps/core.xml"] = strToU8(`${XML_DECL}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title>${esc(title)}</dc:title><dc:creator>ABK Agent</dc:creator>
</cp:coreProperties>`);

  files["docProps/app.xml"] = strToU8(`${XML_DECL}
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Microsoft Office PowerPoint</Application><Slides>${n}</Slides>
</Properties>`);

  return Buffer.from(zipSync(files, { level: 6 }));
}

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
  {
    name: "CreatePresentation",
    description:
      "Create a new PowerPoint presentation (.pptx) from a list of slides. Each slide takes a " +
      "title, optional bullet points and optional speaker notes. Use this to turn an outline, a " +
      "summary or meeting notes into a deck. Creates it in the user's OneDrive by default, or " +
      "directly in a SharePoint site when siteId is given — there is no need to create it in " +
      "OneDrive and then move it. The result is a plain 16:9 deck using the default Office " +
      "theme — content, not visual design.",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "File name, e.g. 'Q3 Review.pptx' ('.pptx' is added if missing)." },
        path: { type: "string", description: "Folder path within the drive, e.g. '/Decks'. Defaults to the root." },
        siteId: { type: "string", description: "Optional. SharePoint site ID to create the deck in. Omit for the user's OneDrive." },
        driveId: { type: "string", description: "Optional. Specific document library, when a SharePoint site has more than one." },
        slides: {
          type: "array",
          description: "The slides, in order.",
          items: {
            type: "object",
            properties: {
              title: { type: "string", description: "Slide title." },
              bullets: { type: "array", items: { type: "string" }, description: "Bullet points for the slide body." },
              notes: { type: "string", description: "Optional speaker notes." },
            },
          },
        },
      },
      required: ["name", "slides"],
    },
  },
  {
    name: "EditPresentation",
    description:
      "Modify an EXISTING PowerPoint presentation in place: add a slide, delete a slide, or " +
      "replace text throughout the deck. A new slide reuses the deck's own layout so it matches " +
      "the existing design. Read the deck first with GetPresentationContent to see the slide " +
      "numbers and exact wording, and confirm with the user before deleting — the file is " +
      "overwritten and a removed slide cannot be recovered from here.",
    inputSchema: {
      type: "object",
      properties: {
        file: { type: "string", description: "The presentation: a path relative to the drive root, or its item ID." },
        siteId: { type: "string", description: "Optional. SharePoint site ID, when the deck lives in a site." },
        action: {
          type: "string",
          enum: ["add_slide", "delete_slide", "replace_text"],
          description: "Which edit to perform.",
        },
        title: { type: "string", description: "add_slide: the new slide's title." },
        bullets: { type: "array", items: { type: "string" }, description: "add_slide: bullet points for the body." },
        notes: { type: "string", description: "add_slide: optional speaker notes." },
        position: {
          type: "number",
          description: "add_slide: 1-based position to insert at. Appended to the end when omitted. delete_slide: which slide to remove.",
        },
        find: { type: "string", description: "replace_text: exact text to look for (case-sensitive)." },
        replace: { type: "string", description: "replace_text: text to put in its place." },
      },
      required: ["file", "action"],
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

// ── Editing an existing package ───────────────────────────────────────────────

/** Adds a slide, reusing the deck's OWN layout so it matches the existing design. */
function addSlideToPackage(files: Parts, spec: SlideSpec, position?: number): number {
  const existing = orderedSlides(files);

  // Part names, relationship IDs and slide IDs must each be unique across the
  // package; derive every one from what is already there rather than assuming
  // a count, since a deck with deletions has gaps.
  const slideNums = Object.keys(files)
    .map((n) => n.match(/^ppt\/slides\/slide(\d+)\.xml$/))
    .filter(Boolean)
    .map((m) => Number(m![1]));
  const num = nextNumericId(slideNums, 1);

  const relsXml = strFromU8(files["ppt/_rels/presentation.xml.rels"]);
  const rIdNums = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  const rId = `rId${nextNumericId(rIdNums, 1)}`;

  const presXml = strFromU8(files["ppt/presentation.xml"]);
  const sldIds = [...presXml.matchAll(/<p:sldId\b[^>]*id="(\d+)"/g)].map((m) => Number(m[1]));
  const sldId = nextNumericId(sldIds, 256);

  // Inherit the layout an existing slide points at; a freshly invented layout
  // would look nothing like the rest of the deck.
  let layoutTarget = "../slideLayouts/slideLayout1.xml";
  if (existing.length) {
    const firstRels = files[`ppt/slides/_rels/${existing[0].part.split("/").pop()}.rels`];
    if (firstRels) {
      const m = strFromU8(firstRels).match(/Target="([^"]*slideLayouts\/[^"]+)"/);
      if (m) layoutTarget = m[1];
    }
  }

  files[`ppt/slides/slide${num}.xml`] = strToU8(slideXml(spec));
  const rels = [`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutTarget}"/>`];
  if (spec.notes) {
    rels.push(`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${num}.xml"/>`);
    files[`ppt/notesSlides/notesSlide${num}.xml`] = strToU8(notesXml(spec.notes));
    files[`ppt/notesSlides/_rels/notesSlide${num}.xml.rels`] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${num}.xml"/>
</Relationships>`);
  }
  files[`ppt/slides/_rels/slide${num}.xml.rels`] = strToU8(`${XML_DECL}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`);

  files["ppt/_rels/presentation.xml.rels"] = strToU8(relsXml.replace(
    "</Relationships>",
    `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${num}.xml"/></Relationships>`
  ));

  const entry = `<p:sldId id="${sldId}" r:id="${rId}"/>`;
  const at = position && position >= 1 && position <= existing.length ? position - 1 : existing.length;
  const updatedList = existing.length === 0
    ? entry
    : (() => {
        const items = [...existing.map((s) => `<p:sldId id="${s.sldId}" r:id="${s.rId}"/>`)];
        items.splice(at, 0, entry);
        return items.join("");
      })();
  files["ppt/presentation.xml"] = strToU8(
    presXml.includes("<p:sldIdLst>")
      ? presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${updatedList}</p:sldIdLst>`)
      : presXml.replace("</p:sldMasterIdLst>", `</p:sldMasterIdLst><p:sldIdLst>${updatedList}</p:sldIdLst>`)
  );

  let ct = strFromU8(files["[Content_Types].xml"]);
  ct = ct.replace("</Types>",
    `<Override PartName="/ppt/slides/slide${num}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    (spec.notes ? `<Override PartName="/ppt/notesSlides/notesSlide${num}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` : "") +
    "</Types>");
  files["[Content_Types].xml"] = strToU8(ct);

  return at + 1;
}

/** Removes a slide by its displayed position. */
function deleteSlideFromPackage(files: Parts, position: number): string {
  const slides = orderedSlides(files);
  const target = slides[position - 1];
  if (!target) throw new Error(`Slide ${position} does not exist — the deck has ${slides.length} slide(s).`);

  const presXml = strFromU8(files["ppt/presentation.xml"]);
  files["ppt/presentation.xml"] = strToU8(
    presXml.replace(new RegExp(`<p:sldId\\b[^>]*r:id="${target.rId}"[^>]*/>`), "")
  );

  const relsXml = strFromU8(files["ppt/_rels/presentation.xml.rels"]);
  files["ppt/_rels/presentation.xml.rels"] = strToU8(
    relsXml.replace(new RegExp(`<Relationship\\b[^>]*Id="${target.rId}"[^>]*/>`), "")
  );

  const base = target.part.split("/").pop()!;
  const slideRels = files[`ppt/slides/_rels/${base}.rels`];
  // Any notes part belongs to this slide alone, so it goes with it.
  const notesTarget = slideRels
    ? strFromU8(slideRels).match(/Target="\.\.\/(notesSlides\/[^"]+)"/)?.[1]
    : undefined;

  const drop = [target.part, `ppt/slides/_rels/${base}.rels`];
  if (notesTarget) drop.push(`ppt/${notesTarget}`, `ppt/notesSlides/_rels/${notesTarget.split("/").pop()}.rels`);
  for (const p of drop) delete files[p];

  // A dangling Override for a part that no longer exists makes the file
  // unopenable, so the content types must be pruned too.
  let ct = strFromU8(files["[Content_Types].xml"]);
  for (const p of drop.filter((x) => x.endsWith(".xml"))) {
    ct = ct.replace(new RegExp(`<Override\\b[^>]*PartName="/${p.replace(/\//g, "\\/")}"[^>]*/>`), "");
  }
  files["[Content_Types].xml"] = strToU8(ct);

  return base;
}

/** Literal find/replace across every slide's visible text and speaker notes. */
function replaceTextInPackage(files: Parts, find: string, replace: string): number {
  const needle = esc(find);
  const replacement = esc(replace);
  let count = 0;
  for (const name of Object.keys(files)) {
    if (!/^ppt\/(slides|notesSlides)\/[^/]+\.xml$/.test(name)) continue;
    const xml = strFromU8(files[name]);
    const updated = xml.replace(/(<a:t(?:\s[^>]*)?>)([^<]*)(<\/a:t>)/g, (whole, open, text, close) => {
      if (!text.includes(needle)) return whole;
      count += text.split(needle).length - 1;
      return open + text.split(needle).join(replacement) + close;
    });
    if (updated !== xml) files[name] = strToU8(updated);
  }
  return count;
}

export class PowerPointToolHandler {
  private token: string;

  constructor(token: string) {
    this.token = token;
  }

  private async createPresentation(
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
    const raw = String(args.name ?? "").trim();
    if (!raw) return { content: [{ type: "text", text: "Error: 'name' is required." }], isError: true };
    const fileName = /\.pptx$/i.test(raw) ? raw : `${raw}.pptx`;

    const slidesArg = Array.isArray(args.slides) ? args.slides : [];
    const slides: SlideSpec[] = slidesArg.map((s) => {
      const o = (s ?? {}) as Record<string, unknown>;
      return {
        title: o.title ? String(o.title) : "",
        bullets: Array.isArray(o.bullets) ? o.bullets.map(String) : [],
        notes: o.notes ? String(o.notes) : undefined,
      };
    });
    if (slides.length === 0) {
      return { content: [{ type: "text", text: "Error: at least one slide is required." }], isError: true };
    }

    const buffer = buildPptx(slides, fileName.replace(/\.pptx$/i, ""));

    const uploadPath = buildUploadPath(
      fileName,
      String(args.path ?? ""),
      args.siteId as string | undefined,
      args.driveId as string | undefined
    );

    const res = await fetch(`https://graph.microsoft.com/v1.0${uploadPath}`, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${this.token}`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      },
      body: new Uint8Array(buffer),
    });
    if (!res.ok) {
      return { content: [{ type: "text", text: `Error creating presentation (${res.status}): ${await res.text()}` }], isError: true };
    }
    const data = await res.json() as { webUrl?: string };
    log(`Created presentation '${fileName}' with ${slides.length} slide(s)`);
    return {
      content: [{ type: "text", text: `✅ Created '${fileName}' with ${slides.length} slide(s).${data.webUrl ? `\nURL: ${data.webUrl}` : ""}` }],
    };
  }

  private async editPresentation(
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
    const err = (text: string) => ({ content: [{ type: "text", text }], isError: true });

    const ref = String(args.file ?? "").trim();
    if (!ref) return err("Error: 'file' is required.");
    const action = String(args.action ?? "");
    const siteId = args.siteId as string | undefined;

    const item = await fetchDriveItem(this.token, ref, siteId);
    if (!/\.pptx$/i.test(item.name)) {
      return err(`'${item.name}' is not a .pptx file.`);
    }
    const files = unzipSync(item.buffer) as Parts;
    const before = orderedSlides(files).length;

    let summary: string;
    if (action === "add_slide") {
      const spec: SlideSpec = {
        title: args.title ? String(args.title) : "",
        bullets: Array.isArray(args.bullets) ? args.bullets.map(String) : [],
        notes: args.notes ? String(args.notes) : undefined,
      };
      const at = addSlideToPackage(files, spec, args.position ? Number(args.position) : undefined);
      summary = `✅ Added slide at position ${at} of ${before + 1}.`;
    } else if (action === "delete_slide") {
      const pos = Number(args.position);
      if (!pos || pos < 1) return err("Error: 'position' (the slide number to delete) is required.");
      if (before <= 1) return err("Error: a presentation must keep at least one slide.");
      deleteSlideFromPackage(files, pos);
      summary = `✅ Deleted slide ${pos}. The deck now has ${before - 1} slide(s).`;
    } else if (action === "replace_text") {
      const find = String(args.find ?? "");
      if (!find) return err("Error: 'find' is required for replace_text.");
      const n = replaceTextInPackage(files, find, String(args.replace ?? ""));
      if (n === 0) {
        return err(`No occurrences of "${find}" were found — the deck was not modified. Matching is case-sensitive.`);
      }
      summary = `✅ Replaced ${n} occurrence(s) of "${find}".`;
    } else {
      return err(`Error: unknown action '${action}'. Use add_slide, delete_slide or replace_text.`);
    }

    const rebuilt = Buffer.from(zipSync(files, { level: 6 }));
    const base = siteId ? `/sites/${siteId}/drive` : "/me/drive";
    const looksLikePath = ref.includes("/") || ref.includes(".");
    const clean = ref.replace(/^\/+|\/+$/g, "");
    const itemPath = looksLikePath ? `${base}/root:/${encodeURI(clean)}` : `${base}/items/${ref}`;

    const up = await fetch(`https://graph.microsoft.com/v1.0${itemPath}/content`, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${this.token}`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      },
      body: new Uint8Array(rebuilt),
    });
    if (!up.ok) return err(`Error saving '${item.name}' (${up.status}): ${await up.text()}`);

    log(`Edited '${item.name}': ${action}`);
    return { content: [{ type: "text", text: `${summary} Saved to '${item.name}'.` }] };
  }

  async handleToolCall(
    toolName: string,
    args: Record<string, unknown>
  ): Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> {
    try {
      if (toolName === "CreatePresentation") {
        return this.createPresentation(args);
      }
      if (toolName === "EditPresentation") {
        return this.editPresentation(args);
      }
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
      // Read the order from <p:sldIdLst>, falling back to file names only if
      // that is unreadable. Sorting by file name is wrong for any deck whose
      // slides have been inserted, reordered or deleted — and the numbers
      // reported here are what EditPresentation is then asked to act on, so
      // getting this wrong would edit the wrong slide.
      const ordered = orderedSlides(files).map((s) => s.part).filter((p) => files[p]);
      const slideNames = ordered.length
        ? ordered
        : Object.keys(files)
            .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
            .sort((a, b) => slideNumber(a) - slideNumber(b));

      if (slideNames.length === 0) {
        return { content: [{ type: "text", text: `'${item.name}' contains no slides.` }], isError: true };
      }

      const sections: string[] = [];
      for (let i = 0; i < slideNames.length; i++) {
        const name = slideNames[i];
        // Position in the deck, which is what the user and EditPresentation
        // both mean by "slide 3" — not the number embedded in the file name.
        const n = i + 1;
        const body = decodeXmlEntities(extractPartText(strFromU8(files[name])));

        let notes = "";
        if (includeNotes) {
          // The notes part is found through the slide's own relationships, not
          // by number: after an insert, slide 2 of the deck may be slide5.xml
          // and its notes notesSlide5.xml.
          const baseName = name.split("/").pop();
          const relPart = files[`ppt/slides/_rels/${baseName}.rels`];
          const notesTarget = relPart
            ? strFromU8(relPart).match(/Target="\.\.\/(notesSlides\/[^"]+)"/)?.[1]
            : undefined;
          const notesPart = notesTarget
            ? files[`ppt/${notesTarget}`]
            : files[`ppt/notesSlides/notesSlide${slideNumber(name)}.xml`];
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
