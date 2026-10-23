#!/usr/bin/env node
/*
 * lilt-docx.js  v1.0.0  (21 Oct 2026)
 * Builds a submission document in Lilt standard format from a JSON spec.
 *
 *   node lilt-docx.js spec.json "Output File Name.docx"
 *
 * Format (Lilt RFI/RFP Response Guidance section 4; Lilt brand context):
 *   DM Sans throughout; LILT Navy #0B1A28 text and headings; Teal #36BBA1 accent;
 *   navy LILT logo top left on every page; title light and large, left aligned;
 *   BLUF as a tinted callout with a teal bar (the one sanctioned visual device);
 *   navy table header row; identity block and page number in the footer;
 *   no decorative rules.  Text is used exactly as given: two spaces after periods
 *   and the identity block are preserved, never reflowed.
 *
 * Spec:
 *   { "title": "...", "label": "...", "meta": "...", "identity": "...",
 *     "blocks": [
 *       {"type":"h1","text":"..."}, {"type":"h2","text":"..."},
 *       {"type":"p","text":"... **bold** ..."},
 *       {"type":"bluf","text":"..."},
 *       {"type":"bullets","items":["**Keyword.**  Explanation", ...]},
 *       {"type":"table","header":["A","B"],"rows":[["..",".."]],"widths":[0.3,0.7]},
 *       {"type":"pagebreak"} ] }
 *   **double asterisks** mark bold.  Requires the docx package (npm install docx).
 */
const fs = require("fs");
const d = require("docx");
const { Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer, Table, TableRow,
  TableCell, WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, PageNumber,
  TabStopType, HeadingLevel, PageBreak } = d;

const NAVY = "0B1A28", TEAL = "36BBA1", MUTED = "5B6670", RULE = "C9CED6", TINT = "E8F6F2";
const FONT = "DM Sans";
const PAGE_W = 12240, MARGIN = 1296;               // US Letter, 0.9 inch margins
const CONTENT_W = PAGE_W - 2 * MARGIN;

const LOGO_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAOwAAABXCAYAAAAH3qmiAAAUpklEQVR42u1da2xlV3ld69vnXj8nITOZBE+mGIqVmRiaCCxRGpBGULWK6A+kVkYgWtQf/VNVrYBCKwEqigoVVVWVUPoDVaJqhVTK0KoVEoVGtBraKunDvNI4k4lnooHJeGY8D3v8vvecvfpj7/uwY3v8uNe+Y+8lHdnXj3Pv3mev/b33R2wNBGDAsAOOehwonBEAAThg4064G+HSFLxqPpimIWG3cadFZ80Spfu1jwya8jebc4OiLwNOgLh/hY4BBQxERYaXi0r1xeWp8+ejpK3Nj5peJyTsOmGJU6cczkwZMF7F8bd3H/K3fsGDvwTpMYG9gHoINksZ7eP5cYKqJKYFzFM874VvLD7U9w2MjVWBUxnwgIDTAFCkJZWw+4QdGipjYmK576GHH5Xs9yk8LqhC8raAJUgFBYv/r309Pw4OnhIhQk5iH4A+Es9Zjs/NXjv7bJS0SPZtwh4QdtQBp4u+Yw+/D+JHRfRRvA1CkExgWLuNRbq/CQsSUFR76cNgRQK9Ag3U5xcuv/jlJjs/SdmEtjpPVhlt47732Ilfh/hZQDOgzYIwQCIJQgfM4cJmtd8IiqRALAGsgnhXuf/+UnXuxjMYGckwOZns2YS2E5YYGurCzZt510PD7zbgMwBvAfSBrACFEhp2qwR6kJ7h9X4lcBhnFLVRq2CcNweyCnCZ4Cl36MjN/NxzP0TDEZWQ0EYJ+4Y3WI/1HTOPLzCof/Mkuik4QgbQwCYnE1cs5H0sccPwGuMka8MlUAK5TMqM/NnuQ/efqcxev44U8kloE6z+dWysSrgPETgWVD2VJDXbqILka44VSo6S7fPFSUIumgEQUAjygOrOJUrdAhcldHnhN+KPs0TahPYQdmSkBMD3Dw0dhefPA5wD5SAagKWGw0VFJKuaRI/tB7V3AxWWIK1hxzI4nshCZB7mg1n81TSgd/Y9+MY3AciBUdsn85PQUYQd6xcA+EV7gsTrBS0KWK6rfA1JQaxUf/eDnSaBEqhVG1FtfD5oFUHVIERCRoUr/k2FUgaAIu6jld4Rfj6cEioS2qESnwmLyvMtkEoEjUI52q1uDbKuJi7v8mu1wdokeqk1tIr6mFcykgahItPjIZniSZ8Im9BqZAAKYKQEzB0DeA1EFSAFeEBig7T1VbxPfElqUocNoCCEVEvGeOsdxtrIGgkqM8mKvI7jp+f6cAEz2P+JJQl7QFjg8EyPITsholtQN8QMVEGxKsBCvjD318JTfUyRmzUHGn3Yp7gEYhGiC+Nfl/cEo40riER339xy9zwws6vbz/7aGNJGt86cBG/mzWrFD2RfBZRFyeOCfVfP2tmPE7h6TIHAlEHMKZwU8FZA87HCoWazbjAXKgSUXKm4B8DVFs4bN3jv6Pw7ReBMvk/IyBU63e5+ts70zwwPlzA+XiUGBnoxObmMlFK3EoOD3X15z6ckvZfANMAypAJrhWtIJ6CC4IiynPxQ5fLZc1hV7bR9P8OIA8aKuJFW1/gb10Tou+U5tmJuWr2BdGr1lYtzZRbJ6sOHHXXxQx/ca3TUYWioCxcvLqGc/TmEKxDKiBlP6zzwprBPSx82MThYjmT1a5P109ZE0s4vPhiF2wRZGRx3m6rDtlVftyNR0TR/nahJFrXxZSt35AsOQH6g7YfTpxmJwd45Pz9fwhSAw5KW2FhAK+dHNemmsFuzJfa+AfA9ld4neOzkr1LRiUUWgGdjrX1N0MkMwPJ8bh/D1Pjc7k/aSAkYq/Y9dPIj8vg5ArPBtKIguSZ6VPUfeg0f0jPzr5x7at11dvx4d6+/8kfAyQcILqxjfnhJ98rw9cVXXvw6tpeoYr3HHv4EYCcpzIJs1lT2Hh4CvMnQRXJunnMfzVbaCmPVpAs3Mrumuu4p+vx0GYIDIQnNudNqtl1rZBbgWam0gLCjBE7DwJ+C/OMwXpVXBsGDLMVsK8UQcRcM1/uzvHsOmNt9n8NYHvx4OAHo7SBvCspChisYJRcBFiTulefUOjZ5eF09bLDKY6ReJ2FhLacfw73uk9ezO1OD+VYCjwGaEeCCM7L2e2ojh+PuWNT0lHoF3Dqy1JdliZ8bWQ5dgpcpuoG1/gNs265MKhc4LXEaRAmAJ1gKxK0HnkoQZsg99uRT8wBvCZwOxSL1hechGIgCgiexsRbgc9IwC/AWQ1XUWnNbUCQMizv80AsQZkTOMmgwYVbrYT/uqbQN4VUuI2hYyFYatqOIJydsaIRE3RHJUbUrj4yQHAgX1EsSVAg1qR4GdoLcnmtygiHkXbv4tfbz4MkWIdBRfhP2pgyKVVFrO/oAyEVH3w4+s0wh18DQyI2vzW0HqMe15x3yIbKVhu3pTdzgdOJQQsIeIYtGfk9f0fukaF0QqhAdLKp+vqZmKWbzoCoqA3hrYbHyFG5dmEnTmJCwe4S1Q76/1xs/bnUnWRNsLVuYUJEv9pV6vxQzelJmSkLCLhFWVuoqiurSNKT+WJnCje1gEsCMK/vkVU7YN86CbXiFVxeFbOXvN/e/bKpHrxE2m1sqql10sbZzU4QVfNkXear3TLiLMQrgR7VVbU2OMq6gzPqwyKi86WBCrUEYMWwGJWBrzqxa6WY4UKLg9sI64a2qhCU1OOEuxjUSWgA4K2hJ4ZyuVWt6Q4lLeh0CUYhcJtAT4+JNUpQ+xPC1SKIHoQY7j/e9sykZQtkGQKTTjuKwtETYhLtV/QWAM7ny4U96py4wrmW/mqCE5Ek21rrkgwPWnM9Y/B6Ax8N53XWVupaU4RXCU4cAPpfnxZ+A5ldzlDSt9R61n0OFoUu8caO8nBInEg40FqbGr+zk//uOnZiTUAqVb1qlEosMpJUR1YVrL13Y6edNNmjCQcd2TyoJCR2iAQy2KZGtJCwVjltiPLXv07XEDNv8+wx1hfc6lQErEyc6EbWyorWt6NQaI6Fl6vGW4QFIVBYyuegAFK+6XaiyLuRZAE82n3KySUwsR/W94yUsQ1H2ihrFpuuUIbV9TNj7VVo/o7tBw1oppkigJKIL9CW0IFehwyVsPEFhcLD7XqnH570GAIdLi/MXL55ZSqslYc/hGQ7eBIrgZFrVKUICYo+bVrxdJxK2JvVd78DJd4kctareUoVeF3ctP1XtudF77OT3BX1t0eafxqVLy0h9WhP2RqPeKDQjkQWkPNQx7zwjsFMIGw3sUxZ7rRa9x058wmh/iNg7jmo6YZR4kLRhSB/s9f1PLQAfbrJ5UwVRwh6sXxKAsX7wPqRXC5EdC5ROsWFrcbECR5/v6X3okS8S/APJS77IG4d5x0vy8nkueRH6zd6BE3/b98CbH0TTURoJCXsgdO70Pe52whKNPjS65/jwfb2l4p+M9lsN6c8Mr3aDW/w5IThz2fvl8n/pfu0jg6ifT5WQsP/QCQs7etRGrSiKz5PZu+Xz6qbVB1IqimVz9qiZ/woGB7vbsbMlJCTC1k/dh/oGnn8/zf0aVOSrpOqd7XCiS97nZPbO3kr3R5OUTUiEbR+KgYGRXrH4JEJ/y+22sDSoEIjf6b3/5EC0Z5OUTUiEbSEcAE1r7j2keyQeAWjbH4u8mXuQGT5Ql74JCYmwrcIpRmPzHaHru1qRaihQ70LnnCyfkLBfCHsmhGGIn0EjlrUDkJIocBADA91JLU5IhG0dQtbH0FCJ5LF4Yid3fE8BoB64t3R/V3q8CYmw7SNvQkJCBxM2nGszMVGVdImt4ayi3J6aqV5fTo83IRG2pThVq3d9Di1xOsmTFMGXMTm5iEYbxoSERNid40zofm74TwWn004/T63Nwr+hkZSRkJAI2yIUAHiE931TKsYZek5tV8p6gCZfXFGOr0YJnqRrQiJsi5FduvTsIoXPxrBOramutkxYOkr8wsL1s5OhOfUDibAJibAtRg6A85OP/p188TegZZAKbFzX2nxcTBVSheYyKf/3hfLin4VxnfbA6VQbm5AI22LE09ZPe+fcR+SL79BltZK79ap2YuW+CgCkZWX5/Hse2Qdx8eJSE6kTEhJh2yBhBYC3L43fXJh074EvngLpSCtFcnpAeeMK1Tiky0DLvPdfpS+eWLr8/E+QUhITEmF38/O8qZi7fPbD8PpFyX8TQgWk0VxGyzJaloFmgCpeepoo3rcw+cIH5q5MTCWyJux3dFI1SyTaaQKw+cmzTwN4uv+1b3rEq3ibPB4EfQ+gHOJV0P/vwuUXf9CkIjORNSERdvdRcyY5AH7uyvMvAHhhAw3BmtTqhISkEu8Roof3VIZ6q4Lhcvg66prU3zw9xoQkYTtG2sbDxEOrgk7r9K47vN4F1E+dZ5tKKDptzjf6nEnCdqC6vOuEIODV1Piz9guB4aDo8LIQme/inIa0TsoLKgBm8rk1mQmtuByODvehI1uiCCKcyYK2NTxswEipxePvOH6kI1R2tmhIwUCGZA8RJKhyebcWt1HIPC0jvSv3YRmtdrxNjc81Xoy6kJDSKRKX8oxhvvHxSlKJDzqK8wT6nIAMgiNkQfEMRQqUgmda0VFGSIolfrujAgoATJLIrsX50sd6Bk54AzOAnlQhiTtiRNQgJE4fcc/85aVLWOyUvRLQEsl39A48fK/APpIFJVP46kJHORWb3WBIyocNjxQz0M/Ml5b+oikZJxG2o+HeKPirBYFcVCHQUyCgWtme6voZSUhGgiR3SwIFldyrCtJI/8sEukMSmLwHcnBn5h0FA9QP6Pzy8j1f7hy7VgZgicJJgG+racmxINoLcKScwE2fcR3bp1OQF1QicOVe3fPlGWCpU2z5RNiNUJkzkCbIEaQEqmn5CxQIUXAQSMAglIBYOz8y4jA21s7YsKJizFhLfCP2KYWoAmBBqBQbNm1XkBUUpwlMd+pTAnG1qXDEx6asDqBtuWmcZAxaRQniNGgd5XBLhN1QwpYE7wHCxQ3WCBhYV4kLicVqApWcy5YBYGys3SEnI2LXbwVVDgznOlPIQy+ielLJdhXvIij/kj9cEFMd9YRqJkoWOj5CkbRxzNvhGiPHmXUiPxJhN8Klao4BToO4DmiJoAdBBeIKpACZxFLc2T2p21Xvq7uoOjanY9bU9Foj7Hq/0h0YyhbVz05FQ+EBLWyuUGigVusgtxUTRSTpJRSEfKcpFomwG6maGKuqdPKTWABEF37Wu+IP4gtP0AQFE3e5jJsIyR6Mp2rskLSMnfvqalvDdpaqAn3sAi4Irm6KEaK03U4KgQwCQRoVtq3Niajmb+tSXiAUHXRbICK1wfRJkhDGL0gWbG4ozAkKStkWNxwGS5hbmTdijWWxzt/sHWHl/b4PVi/8+Ozk9nf+Mzsg6unainQiPYPjowuQE5hHj7VBMSNMsEaX7yhRgpG95Q0jmOUUEckuFCJk5u5wHxoEFxRShjimkIMsQapGiec35ZQjSqGjefww60vJ+lgRWqEXTb9xmx9/3OQaElab0qkNgFiOd6hANEIZao6DsIFUg9caCLHisepeEPagHE26nXG2shs8KZhAY1T3GNS9GF6S6jxVKwetmoJY87Lm3hd3mAsvkD6SpkKh2ZsuBI+Y86Lb+DZFJA+iVx5bdx5tX7FqkrA04DUALm7CbIhsD8H4YKZIFklcFWsptDtzQu7INpH8QSCttnG1AKM1uhZxVVjDqRKlhuoxxnZdvm4fbymeSwEoROYiqpCqInPVT8Xc9L1qWsMuXzWpLN3JhqVXEWO9DaVG8tFsCY3IWd/4iY1PUmmfhCXgzWUeLd/b94VEVgtvJgElSj0Cu0jlIGPxQ7vjvSKBsoQSydk7DkvMQPQR6A/qsAiwCiqjguQRUA4hl401FwFdAPpYi3Hv3v5MSAawJKK8BR6VqejwC1GFoF6DZYVxtKQTxXYIK4Qk2q6islgC4I4ff3v50qVnl3H3JIq3D0NDJUxMVHY+D6c9AJjLv53n2YQnZzPPEhxr9hbySNisrZsPBY+SUfM3brjldTZoASAdvgTh7+mLXGaOUg6ZAb7I4QwuhytYptlVrF+/bLjhlv1A/ilnLNMs37XcMQA5QBSFB62guXzm9a+fw8WL640ZlP5Khm8xRgZyZMgsbKaSLJBWRG7XWsEPAsCRIycOLXbxFYKHmrx661oYQWtw5r3/7cXJs19MQnVNU8O3hjCfJvCk36dztJYJkTb9FhM2kpYUNAPhj+XsX1n4vH5P8oCd/FBTTSuUt+w1PPT85OTYAmIR/g4XYPQznDLgTNO8rn7dpvUxMmIYGysaz/2O68lWjHdkxOHy5RImJ0N63/Cww/h4vgl7v1Niv3cc8/DwcKmnp0djY3WnkhoK11BpYmIib/IJ7Jywhw8P3bPUnV3aJGEbcTWQIdRW5E0q1AHb89T0DQlg2bn80ds/mTgfCXu3H7Xaaol30CRoKzbtlTbszaNOvbNYBHEI2hzJ6/4Qycc0roOspDR971FU/Owmd+e7akfq0Pt1Olq6YYfA+IsvzhM6G48s8FtcratjbQf18iEEwavlClLnvIS2Gf6GEKB+Lmh02lbGdLrgGQpi//vWrQszr7LlEhJaQ9hTIYPK9A/yXrXyrIRtaivkP4eXo6lJdUJbDDBidNRw7Rr7XrryX6S9JSRUp1aNW7NTaID/sTP31tuXxm8ihScS2iQVhAsXDGfO5BQ/Ez2dSottizsfSXr7XCRraiSd0BYE9XdyEsCoVea+ezY7dGSItMdqjaZw8OI0W4Rymsu899+Z788/jps302aX0FbC1mzWDJgquvoPfxfAu0k7HkmbVOMNyBraY/qLhdevFD956XpShRParRLHiozxCgCbmzx3PSv8++T990HL4u9fdQzKgbdZpSotyySdF/17l6+eezmei5v6+yS02YZtXoiAzVw99zLhfxG++GvQDDRXlygHl7CK8yPQHMyV5ItvZ1Z9YuGVl34IjJQwNJO0kYRdsGGbMTpq6OnJKuf+b646d+Mfu/qP/AjQMYCDJHdy3MjdDZKkxUO/8AOJn1i4/MLHlm/fuhnMiUmPm28gMJk0kIRdIyxx9KjhmWeK0HTqoipz1194dHbwK9f6q9+jsCTASN13wOK1uaCzkL4lzz9dcLO/W33l/P/E+csaZsVkUocT2or/BzVU5lg/AAslAAAAAElFTkSuQmCC", "base64");

function runs(text, base = {}) {
  const out = [];
  String(text).split(/(\*\*[^*]+\*\*)/).forEach(part => {
    if (!part) return;
    const bold = part.startsWith("**") && part.endsWith("**");
    out.push(new TextRun({ text: bold ? part.slice(2, -2) : part, bold: bold || base.bold,
      color: base.color || NAVY, size: base.size, font: FONT, allCaps: base.allCaps,
      characterSpacing: base.characterSpacing }));
  });
  return out;
}

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: RULE };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

function table(b) {
  const n = b.header.length;
  const fr = b.widths || Array(n).fill(1 / n);
  const sum = fr.reduce((a, c) => a + c, 0);
  const widths = fr.map(f => Math.floor(CONTENT_W * f / sum));
  widths[n - 1] += CONTENT_W - widths.reduce((a, c) => a + c, 0);
  const cell = (text, i, head) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders,
    shading: head ? { type: ShadingType.CLEAR, fill: NAVY, color: "auto" } : undefined,
    margins: { top: 90, bottom: 90, left: 120, right: 120 },
    children: [new Paragraph({ spacing: { before: 0, after: 0, line: 260 },
      children: runs(text, { size: 18, bold: head || (b.boldFirstColumn && i === 0),
        color: head ? "FFFFFF" : NAVY }) })] });
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: b.header.map((t, i) => cell(t, i, true)) }),
      ...b.rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, i, false)) }))] });
}

function block(b) {
  switch (b.type) {
    case "h1": return [new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true,
      spacing: { before: 240, after: 120 }, children: runs(b.text, { bold: true, size: 28 }) })];
    case "h2": return [new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true,
      spacing: { before: 200, after: 80 }, children: runs(b.text, { bold: true, size: 22 }) })];
    case "p": return [new Paragraph({ spacing: { before: 0, after: 140, line: 288 },
      children: runs(b.text, { size: 21 }) })];
    case "bluf": return [new Paragraph({ spacing: { before: 60, after: 200, line: 288 },
      shading: { type: ShadingType.CLEAR, fill: TINT, color: "auto" },
      border: { left: { style: BorderStyle.SINGLE, size: 36, color: TEAL, space: 10 } },
      indent: { left: 200, right: 120 },
      children: [...runs("**Bottom line.**  ", { size: 21 }), ...runs(b.text, { size: 21 })] })];
    case "bullets": return b.items.map(t => new Paragraph({ numbering: { reference: "bullets", level: 0 },
      spacing: { before: 0, after: 80, line: 276 }, children: runs(t, { size: 21 }) }));
    case "table": return [table(b), new Paragraph({ spacing: { before: 0, after: 160 }, children: [] })];
    case "pagebreak": return [new Paragraph({ children: [new PageBreak()] })];
    default: throw new Error("unknown block type: " + b.type);
  }
}

function build(spec, outPath) {
  const header = new Header({ children: [new Paragraph({ spacing: { after: 240 },
    children: [new ImageRun({ type: "png", data: LOGO_PNG, transformation: { width: 92, height: 34 } })] })] });
  const footer = new Footer({ children: [
    new Paragraph({ style: "LiltFooter", spacing: { before: 0, after: 40 }, children: runs(spec.identity || "", { size: 15, color: MUTED }) }),
    new Paragraph({ style: "LiltFooter", alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: "Page ", size: 15, color: MUTED, font: FONT }),
        new TextRun({ children: [PageNumber.CURRENT], size: 15, color: MUTED, font: FONT }),
        new TextRun({ text: " of ", size: 15, color: MUTED, font: FONT }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 15, color: MUTED, font: FONT })] })] });
  const top = [];
  if (spec.label) top.push(new Paragraph({ spacing: { before: 120, after: 60 },
    children: runs(spec.label, { size: 18, bold: true, color: TEAL, allCaps: true, characterSpacing: 20 }) }));
  top.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { before: 0, after: 80 },
    children: runs(spec.title, { size: 40 }) }));
  if (spec.meta) top.push(new Paragraph({ spacing: { before: 0, after: 240 },
    children: runs(spec.meta, { size: 19, color: MUTED }) }));
  const doc = new Document({
    creator: "Lilt, Inc.", title: spec.title,
    styles: { default: { document: { run: { font: FONT, color: NAVY, size: 21 } } },
      paragraphStyles: [
        { id: "Title", name: "Title", basedOn: "Normal", run: { font: FONT, size: 40, color: NAVY } },
        { id: "LiltFooter", name: "Lilt Footer", basedOn: "Normal", run: { font: FONT, size: 15, color: MUTED } },
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 28, bold: true, color: NAVY }, paragraph: { outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 22, bold: true, color: NAVY }, paragraph: { outlineLevel: 1 } }] },
    numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET,
      text: "\u2022", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 360, hanging: 240 } }, run: { color: TEAL } } }] }] },
    sections: [{ properties: { page: { size: { width: PAGE_W, height: 15840 },
      margin: { top: 1152, bottom: 1152, left: MARGIN, right: MARGIN, header: 576, footer: 432 } } },
      headers: { default: header }, footers: { default: footer },
      children: [...top, ...spec.blocks.flatMap(block)] }] });
  return Packer.toBuffer(doc).then(buf => { fs.writeFileSync(outPath, buf); return buf.length; });
}

if (require.main === module) {
  const [specPath, outPath] = process.argv.slice(2);
  if (!specPath || !outPath) { console.error("usage: node lilt-docx.js spec.json out.docx"); process.exit(2); }
  const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
  for (const k of ["title", "blocks"]) if (!spec[k]) { console.error("SPEC INVALID: missing " + k); process.exit(2); }
  build(spec, outPath).then(n => console.log(outPath + " (" + n + " bytes)"))
    .catch(e => { console.error("BUILD FAILED: " + e.message); process.exit(2); });
}
module.exports = { build };
