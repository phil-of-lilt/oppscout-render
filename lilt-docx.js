#!/usr/bin/env node
/*
 * lilt-docx.js  v1.1.1  (23 Oct 2026)
 * Builds a buyer-facing response in Lilt standard format from a JSON spec.
 *
 *   node lilt-docx.js spec.json "Output File Name.docx"
 *
 * v1.1.1: matches Lilt's RFI 13225 (WA DES) response exactly, measured from its XML:
 *   the LILT logo in its own colors; italic date; bold underlined title; run-in identity
 *   lines; contact table with black rules; justified body; headings with a thin teal rule;
 *   the buyer's question restated in italics; callout boxes with a thin black outline and
 *   a teal bar; purple lettered subheads; teal bullets; navy table header with light teal
 *   rows; page 1 footer for a notice, later pages page number plus document title.
 *   Refuses em-dashes in the spec (BLOCKED unless "allowEmDash": true).
 * v1.1.0: first version standardized to the WA DES layout.
 *
 * Spec (**double asterisks** mark bold; URLs and email addresses become links):
 *   { "date": "23 October 2026",
 *     "title": "Lilt: Response to ...",
 *     "identity": [["Company","Lilt, Inc."],["Address","..."],["UEI","..."],["CAGE","..."],["DUNS","..."]],
 *        (or a list of lines, each a list of pairs, to control where the line breaks)
 *     "contacts": [["Phil Stiefel","LILT - Head of Public Sector","phil@lilt.com, (949) 796-3247"]],
 *     "footer_title": "LILT - Response to ...",
 *     "notice": {"title":"Public Records Notice","text":"..."},      (optional, page 1 footer)
 *     "justify": true,                                                (default true)
 *     "blocks": [
 *       {"type":"h1","text":"Q1.  Implementation Approach"},
 *       {"type":"question","text":"RFI Q1: Please describe ..."},
 *       {"type":"bluf","text":"..."}            (label defaults to "Bottom line up front.";
 *       {"type":"bluf","label":"Lilt recommends:","text":"..."}   any label may be given)
 *       {"type":"p","text":"..."},
 *       {"type":"h2","text":"a.  Licensing model"},
 *       {"type":"leads","items":["**Phase 1: Kickoff.**  ...", "..."]},   (indented paragraphs)
 *       {"type":"bullets","items":["**Keyword.**  ..."]},
 *       {"type":"table","header":["A","B"],"rows":[["..",".."]],"widths":[0.3,0.7],"boldFirstColumn":true},
 *       {"type":"pagebreak"} ] }
 *   v1.0.0 specs still build: "label" prints above the title, "meta" is used as the date
 *   line, and a string "identity" prints as one line.  Requires the docx package.
 */
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, ImageRun, Footer, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, PageNumber, HeadingLevel,
  PageBreak, ExternalHyperlink, UnderlineType, LineRuleType } = require("docx");

const NAVY = "0B1A28", TEXT = "1A2730", BLUF_TEXT = "26333D", QUESTION = "30404B", TEAL = "36BBA1",
  PURPLE = "413B95", DATE = "595959", LINK = "0563C1", MAIL = "1155CC", TINT = "EAF6F2",
  HEAD_RULE = "BFBFBF", ROW_RULE = "D6DCE0";
const PAGE_W = 12240, MARGIN = 1080, CONTENT_W = PAGE_W - 2 * MARGIN;
let FONT = "Source Sans Pro", JUSTIFY = true;

const LOGO_PNG = Buffer.from(
  "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", "base64");

const LINKISH = /(https?:\/\/[^\s)]+|[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/;
function textRuns(text, o) {
  return String(text).split(LINKISH).flatMap(seg => {
    if (!seg) return [];
    if (LINKISH.test(seg) && seg.match(LINKISH)[0] === seg) {
      const mail = seg.includes("@") && !seg.startsWith("http");
      return [new ExternalHyperlink({ link: mail ? "mailto:" + seg : seg, children: [new TextRun({ text: seg,
        color: mail ? MAIL : LINK, underline: { type: UnderlineType.SINGLE }, size: o.size, font: FONT,
        italics: o.italics, bold: o.bold })] })];
    }
    return [new TextRun({ text: seg, bold: o.bold, italics: o.italics, color: o.color || TEXT, size: o.size,
      font: FONT, underline: o.underline ? { type: UnderlineType.SINGLE } : undefined,
      allCaps: o.allCaps, characterSpacing: o.characterSpacing })];
  });
}
function runs(text, base = {}) {
  return String(text).split(/(\*\*[^*]+\*\*)/).flatMap(part => {
    if (!part) return [];
    const bold = part.startsWith("**") && part.endsWith("**");
    return textRuns(bold ? part.slice(2, -2) : part, { ...base, bold: bold || base.bold });
  });
}
const rule = (sz, color) => ({ style: BorderStyle.SINGLE, size: sz, color });
const allSides = r => ({ top: r, bottom: r, left: r, right: r });
const gap = after => new Paragraph({ spacing: { before: 0, after }, children: [] });
const align = () => JUSTIFY ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;

function widthsFor(fr, n, total = CONTENT_W) {
  const f = fr || Array(n).fill(1 / n), sum = f.reduce((a, c) => a + c, 0);
  const w = f.map(x => Math.floor(total * x / sum));
  w[n - 1] += total - w.reduce((a, c) => a + c, 0);
  return w;
}
function table(b) {
  const n = b.header.length, w = widthsFor(b.widths, n), boldFirst = b.boldFirstColumn !== false;
  const cell = (t, i, head) => new TableCell({ width: { size: w[i], type: WidthType.DXA },
    borders: allSides(rule(1, head ? HEAD_RULE : ROW_RULE)),
    shading: { type: ShadingType.CLEAR, fill: head ? NAVY : TINT, color: "auto" },
    margins: { top: head ? 80 : 70, bottom: head ? 80 : 70, left: 100, right: 100 },
    children: [new Paragraph({ spacing: { before: 0, after: 0 },
      children: runs(t, { size: 19, bold: head || (boldFirst && i === 0), color: head ? "FFFFFF" : TEXT }) })] });
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: w,
    borders: { ...allSides(rule(4, "000000")), insideHorizontal: rule(4, "000000"), insideVertical: rule(4, "000000") },
    rows: [new TableRow({ tableHeader: true, children: b.header.map((t, i) => cell(t, i, true)) }),
      ...b.rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, false)) }))] }),
    gap(60)];
}
function callout(label, text) {
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    borders: allSides(rule(4, "000000")),
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: TINT, color: "auto" },
      borders: { left: rule(18, TEAL) },
      margins: { top: 90, bottom: 90, left: 150, right: 150 },
      children: [new Paragraph({ spacing: { before: 0, after: 0, line: 260, lineRule: LineRuleType.AUTO },
        children: [...runs(label + " ", { size: 21, bold: true, color: NAVY }), ...runs(text, { size: 21, color: BLUF_TEXT })] })] })] })] }),
    gap(110)];
}
function block(b) {
  switch (b.type) {
    case "h1": return [new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, spacing: { before: 260, after: 110 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: TEAL, space: 4 } },
      children: runs(b.text, { bold: true, size: 24, color: NAVY }) })];
    case "h2": return [new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 180, after: 80 },
      children: runs(b.text, { bold: true, size: 21, color: PURPLE }) })];
    case "question": return [new Paragraph({ keepNext: true, spacing: { before: 120, after: 70, line: 252, lineRule: LineRuleType.AUTO },
      children: runs(b.text, { italics: true, size: 21, color: QUESTION }) })];
    case "bluf": return callout(b.label || "Bottom line up front.", b.text);
    case "p": return [new Paragraph({ alignment: align(), spacing: { before: 0, after: 150, line: 264, lineRule: LineRuleType.AUTO }, children: runs(b.text) })];
    case "leads": return b.items.map(t => new Paragraph({ indent: { left: 460 }, spacing: { before: 0, after: 90, line: 260, lineRule: LineRuleType.AUTO },
      children: runs(t) }));
    case "bullets": return b.items.map(t => new Paragraph({ numbering: { reference: "bullets", level: 0 },
      spacing: { before: 0, after: 90, line: 260, lineRule: LineRuleType.AUTO }, children: runs(t) }));
    case "table": return table(b);
    case "pagebreak": return [new Paragraph({ children: [new PageBreak()] })];
    default: throw new Error("SPEC INVALID: unknown block type " + b.type);
  }
}
function identityLines(id) {
  const lines = Array.isArray(id[0][0]) ? id : [id];
  return lines.map(pairs => new Paragraph({ spacing: { before: 0, after: 0 }, children: pairs.flatMap(([k, v], i) =>
    [...runs((i ? " " : "") + k + ": ", { size: 20, bold: true }), ...runs(v, { size: 20 })]) }));
}
function firstPage(spec) {
  const out = [new Paragraph({ spacing: { before: 0, after: 60 },
    children: [new ImageRun({ type: "png", data: LOGO_PNG, transformation: { width: 156, height: 56 } })] })];
  const date = spec.date || spec.meta;
  if (date) out.push(new Paragraph({ spacing: { before: 0, after: 160 }, children: runs(date, { italics: true, size: 26, color: DATE }) }));
  if (spec.label) out.push(new Paragraph({ spacing: { before: 0, after: 60 },
    children: runs(spec.label, { size: 18, bold: true, color: TEAL, allCaps: true, characterSpacing: 20 }) }));
  out.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { before: 0, after: 80, line: 252, lineRule: LineRuleType.AUTO },
    children: runs(spec.title, { bold: true, size: 28, underline: true }) }));
  if (Array.isArray(spec.identity)) out.push(...identityLines(spec.identity), gap(140));
  else if (spec.identity && !spec.footer_title) out.push(new Paragraph({ spacing: { before: 0, after: 140 }, children: runs(spec.identity, { size: 20 }) }));
  if (spec.contacts && spec.contacts.length) {
    const n = spec.contacts.length, w = widthsFor(null, n, 9840), black = rule(6, "000000");
    out.push(new Table({ width: { size: 9840, type: WidthType.DXA }, columnWidths: w,
      borders: { ...allSides(black), insideHorizontal: black, insideVertical: black },
      rows: [new TableRow({ children: spec.contacts.map((c, i) => new TableCell({ width: { size: w[i], type: WidthType.DXA },
        borders: allSides(black), margins: { top: 100, bottom: 100, left: 100, right: 100 },
        children: c.map(t => new Paragraph({ spacing: { before: 0, after: 0 }, children: runs(t) })) })) })] }));
  }
  return out;
}
function footers(spec) {
  const title = spec.footer_title || (typeof spec.identity === "string" ? spec.identity : spec.title);
  const later = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 20, font: FONT, color: TEXT })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: runs(title, { size: 16, bold: true }) })];
  const first = spec.notice ? [
    new Paragraph({ spacing: { before: 0, after: 0 }, children: runs(spec.notice.title, { size: 16, bold: true }) }),
    new Paragraph({ spacing: { before: 0, after: 0 }, children: runs(spec.notice.text, { size: 16 }) })] : later;
  return { first: new Footer({ children: first }), default: new Footer({ children: later }) };
}
function check(spec) {
  for (const k of ["title", "blocks"]) if (!spec[k]) throw new Error("SPEC INVALID: missing " + k);
  if (!spec.allowEmDash && JSON.stringify(spec).includes("\u2014"))
    throw new Error("SPEC INVALID: em-dash found; rewrite those sentences with commas, colons or parentheses");
}
function build(spec, outPath) {
  try { check(spec); } catch (e) { return Promise.reject(e); }
  if (spec.font) FONT = spec.font;
  JUSTIFY = spec.justify !== false;
  const doc = new Document({ creator: "Lilt, Inc.", title: spec.title,
    styles: { default: { document: { run: { font: FONT, color: TEXT, size: 22 } } },
      paragraphStyles: [
        { id: "Title", name: "Title", basedOn: "Normal", run: { font: FONT, size: 28, bold: true } },
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 24, bold: true, color: NAVY }, paragraph: { outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 21, bold: true, color: PURPLE }, paragraph: { outlineLevel: 1 } }],
      characterStyles: [{ id: "Hyperlink", name: "Hyperlink", run: { color: LINK, underline: { type: UnderlineType.SINGLE } } }] },
    numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022",
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 520, hanging: 250 } }, run: { color: TEAL } } }] }] },
    sections: [{ properties: { titlePage: true, page: { size: { width: PAGE_W, height: 15840 },
      margin: { top: 720, bottom: 1080, left: MARGIN, right: MARGIN, header: 708, footer: 720 } } },
      footers: footers(spec), children: [...firstPage(spec), ...spec.blocks.flatMap(block)] }] });
  return Packer.toBuffer(doc).then(buf => { fs.writeFileSync(outPath, buf); return buf.length; });
}

if (require.main === module) {
  const [specPath, outPath] = process.argv.slice(2);
  if (!specPath || !outPath) { console.error("usage: node lilt-docx.js spec.json out.docx"); process.exit(2); }
  let spec;
  try { spec = JSON.parse(fs.readFileSync(specPath, "utf8")); } catch (e) { console.error("SPEC INVALID: " + e.message); process.exit(2); }
  build(spec, outPath).then(n => console.log(outPath + " (" + n + " bytes)"))
    .catch(e => { console.error(e.message.startsWith("SPEC") ? e.message : "BUILD FAILED: " + e.message); process.exit(2); });
}
module.exports = { build };
