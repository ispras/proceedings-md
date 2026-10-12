import WordDocument from "../word/word-document";
import mdast from "mdast";
import { DocumentJsonMeta } from "../document-json-meta";
import * as XML from "../xml";
import * as OXML from "../word/oxml";
import ParagraphTemplateSubstitution from "../word-templates/paragraph-template-substitution";
const languages = ["ru", "en"];
import InlineTemplateSubstitution from "../word-templates/inline-template-substitution";
import { Num } from "../word/numbering";
import ImageSize from "image-size";
import fs from "fs";
import path from "path";
import { FormulaMetrics, formulaKey } from "../markdown/formula-render";
import DocumentReferences from "../markdown/document-references";
const MAX_BLOCK_WIDTH_CM = 14;
const FIGURE_DPI = 300;
export function normalizeFormula(value: string): string {
  return value.replace(/^`+|`+$/g, "");
}

export interface NumberingStyle {
  numId: string;
  styleId: string;
}

let DecimalNumberingStyle: NumberingStyle = {
  styleId: "ispNumList",
  numId: "33",
};
let BulletNumberingStyle: NumberingStyle = { styleId: "ispList1", numId: "43" };

export class GenerationContext {
  doc: WordDocument;
  meta: DocumentJsonMeta;
  references: DocumentReferences;
  nodeStack: XML.Node[] = [];
  numberingStack: NumberingStyle[] = [];
  formulaMetrics: Map<string, FormulaMetrics> = new Map();
  tableAlign: (string | null)[] = [];
  tableColumn = 0;
  tableRow = 0;
  sourceDir: string;
  imagePaths: Map<string, string> = new Map();
  formulaCounter: number = 0;
  inlineRefMode: "ref" | "cite" | null = null;
  vertAlignStack: ("subscript" | "superscript")[] = [];
  boldDepth: number = 0;
  italicDepth: number = 0;

  visit(element: mdast.Node) {
    const visitor = visitors[element.type];
    if (!visitor) {
      if (["root", "yaml"].includes(element.type)) return;
      throw new Error(`Unsupported Markdown node: ${element.type}`);
    }
    visitor(element as any, this);
  }

  visitChildren(element: mdast.Parent) {
    for (let child of (element as mdast.Parent).children) {
      this.visit(child);
    }
  }

  getCurrentNode() {
    return this.nodeStack[this.nodeStack.length - 1];
  }

  pushNode(node: XML.Node) {
    this.nodeStack.push(node);
  }

  popNode() {
    this.nodeStack.pop();
  }

  pushNumberingStyle(style: NumberingStyle) {
    this.numberingStack.push(style);
  }

  popNumberingStyle() {
    this.numberingStack.pop();
  }

  getNumberingStyle() {
    return this.numberingStack[this.numberingStack.length - 1];
  }
}

const visitors: {
  [K in keyof mdast.RootContentMap]?: (
    source: mdast.RootContentMap[K],
    ctx: GenerationContext,
  ) => void;
} = {
  paragraph: (node, ctx) => {
    const onlyImages =
      node.children.length > 0 &&
      node.children.every((c) => c.type === "image");
    if (onlyImages) {
      node.children.forEach((child) =>
        visitBlockImage(child as mdast.Image, ctx),
      );
      return;
    }

    let numberingStyle = ctx.getNumberingStyle();
    let styleId = numberingStyle
      ? numberingStyle.styleId
      : ctx.doc.styles.resource.getStyleByName("ispText_main").getId();
    let paragraph = OXML.buildParagraphWithStyle(styleId);

    if (numberingStyle) {
      paragraph
        .getChild("w:pPr")
        .pushChild(
          XML.Node.build("w:numPr").appendChildren([
            XML.Node.build("w:ilvl").setAttr(
              "w:val",
              String(ctx.numberingStack.length - 1),
            ),
            XML.Node.build("w:numId").setAttr("w:val", numberingStyle.numId),
          ]),
        );
    }

    ctx.getCurrentNode().pushChild(paragraph);
    ctx.pushNode(paragraph);
    ctx.visitChildren(node);
    ctx.popNode();
  },

  strong: (node, ctx) => {
    ctx.boldDepth++;
    ctx.visitChildren(node);
    ctx.boldDepth--;
  },

  code: (node, ctx) => {
    let styleId = ctx.doc.styles.resource.getStyleByName("ispListing").getId();

    let raw = XML.Node.build("w:r");
    let lines = node.value.split("\n");
    for (let line of lines) {
      if (raw.getChildrenCount() > 0) {
        raw.pushChild(OXML.buildLineBreak());
      }
      raw.pushChild(OXML.buildTextTag(line));
    }

    const paragraph = OXML.buildParagraphWithStyle(styleId);
    paragraph
      .getChild("w:pPr")
      .appendChildren([
        XML.Node.build("w:keepNext"),
        XML.Node.build("w:keepLines"),
      ]);
    ctx.getCurrentNode().pushChild(paragraph.pushChild(raw));
  },

  link: (node, ctx) => {
    ctx.visitChildren(node);
  },

  delete: (node, ctx) => {
    ctx.visitChildren(node);
  },

  emphasis: (node, ctx) => {
    ctx.italicDepth++;
    ctx.visitChildren(node);
    ctx.italicDepth--;
  },

  text: (node, ctx) => {
    if (ctx.inlineRefMode) {
      let id = node.value.trim();
      let replacement =
        ctx.inlineRefMode === "cite"
          ? ctx.references.getCite(id)
          : ctx.references.getReference(id);
      ctx
        .getCurrentNode()
        .pushChild(OXML.buildRawTag(replacement, vertAlignStyles(ctx)));
      return;
    }

    let styles = vertAlignStyles(ctx);
    ctx.getCurrentNode().pushChild(OXML.buildRawTag(node.value, styles));
  },

  table: (node, ctx) => {
    let styleId = ctx.doc.styles.resource.getStyleByName("Table Grid").getId();

    let table = XML.Node.build("w:tbl");
    let tableProperties = XML.Node.build("w:tblPr").appendChildren([
      XML.Node.build("w:tblStyle").setAttr("w:val", styleId),
    ]);
    tableProperties.pushChild(
      XML.Node.build("w:tblW").setAttrs({ "w:w": "7938", "w:type": "dxa" }),
    );
    const columns = node.children[0]?.children.length ?? 1;
    table.pushChild(tableProperties);
    table.pushChild(
      XML.Node.build("w:tblGrid").appendChildren(
        Array.from({ length: columns }, () =>
          XML.Node.build("w:gridCol").setAttr(
            "w:w",
            String(Math.floor(7938 / columns)),
          ),
        ),
      ),
    );
    ctx.tableAlign = node.align ?? [];
    ctx.tableRow = 0;
    ctx.getCurrentNode().pushChild(table);
    ctx.pushNode(table);
    ctx.visitChildren(node);
    ctx.popNode();
  },

  tableRow: (node, ctx) => {
    ctx.tableColumn = 0;
    ctx.tableRow++;
    const properties = XML.Node.build("w:trPr").pushChild(
      XML.Node.build("w:cantSplit"),
    );
    if (ctx.tableRow === 1) properties.pushChild(XML.Node.build("w:tblHeader"));
    let row = XML.Node.build("w:tr").pushChild(properties);
    ctx.getCurrentNode().pushChild(row);
    ctx.pushNode(row);
    ctx.visitChildren(node);
    ctx.popNode();
  },

  tableCell: (node, ctx) => {
    let styleId = ctx.doc.styles.resource
      .getStyleByName("ispText_main")
      .getId();

    let paragraph = OXML.buildParagraphWithStyle(styleId);
    const align = ctx.tableAlign[ctx.tableColumn++] ?? "left";
    paragraph
      .getChild("w:pPr")
      .pushChild(XML.Node.build("w:jc").setAttr("w:val", align));
    let cell = XML.Node.build("w:tc").appendChildren([
      XML.Node.build("w:tcPr").pushChild(
        XML.Node.build("w:tcW").setAttrs({
          "w:w": String(Math.floor(7938 / Math.max(1, ctx.tableAlign.length))),
          "w:type": "dxa",
        }),
      ),
      paragraph,
    ]);

    ctx.getCurrentNode().pushChild(cell);
    ctx.pushNode(paragraph);
    if (ctx.tableRow === 1) ctx.boldDepth++;
    ctx.visitChildren(node);
    if (ctx.tableRow === 1) ctx.boldDepth--;
    ctx.popNode();
  },

  inlineMath: (node, ctx) => {
    let formula = normalizeFormula(node.value);
    let metrics = ctx.formulaMetrics.get(formulaKey(formula, false));

    if (!metrics) {
      throw new Error("Missing rendered formula: " + formula);
    }
    let positionHalfPts = -Math.round(metrics.depth * 2);

    let run = buildDualBlipImage(ctx, {
      pngBuffer: metrics.png,
      svgBuffer: metrics.svg,
      widthPt: metrics.width,
      heightPt: metrics.height,
      positionHalfPts,
    });

    ctx.getCurrentNode().pushChild(run);
  },

  listItem: (node, ctx) => {
    ctx.visitChildren(node);
  },

  image: (node, ctx) => {
    const imagePath =
      ctx.imagePaths.get(path.resolve(ctx.sourceDir, node.url)) ??
      path.resolve(ctx.sourceDir, node.url);
    let imageBuffer = fs.readFileSync(imagePath);
    let file = ImageSize(imageBuffer);

    let size = getImageSize(node, file.width / file.height);
    let svgBuffer = readSvgSibling(imagePath);

    let run = buildDualBlipImage(ctx, {
      pngBuffer: imageBuffer,
      svgBuffer,
      widthPt: (size[0] * 72) / 2.54,
      heightPt: (size[1] * 72) / 2.54,
      pngName: path.basename(node.url),
    });
    ctx.getCurrentNode().pushChild(run);
  },

  blockquote: (node, ctx) => {
    ctx.visitChildren(node);
  },

  break: (node, ctx) => {
    ctx
      .getCurrentNode()
      .pushChild(XML.Node.build("w:r").pushChild(OXML.buildLineBreak()));
  },

  heading: (node, ctx) => {
    let headings = [
      "ispSubHeader-1 level",
      "ispSubHeader-2 level",
      "ispSubHeader-3 level",
    ];

    let level = node.depth;

    if (level > headings.length) {
      level = headings.length as typeof node.depth;
    }

    let styleId = ctx.doc.styles.resource
      .getStyleByName(headings[level - 1])
      .getId();
    let paragraph = OXML.buildParagraphWithStyle(styleId);

    paragraph
      .getChild("w:pPr")
      .pushChild(
        XML.Node.build("w:numPr").appendChildren([
          XML.Node.build("w:ilvl").setAttr("w:val", "0"),
          XML.Node.build("w:numId").setAttr("w:val", "0"),
        ]),
      );

    ctx.getCurrentNode().pushChild(paragraph);
    ctx.pushNode(paragraph);
    ctx.visitChildren(node);
    ctx.popNode();
  },

  list: (node, ctx) => {
    let numberingStyle = node.ordered
      ? DecimalNumberingStyle
      : BulletNumberingStyle;

    let numId = createNumbering(
      ctx.doc,
      numberingStyle.numId,
      ctx.numberingStack.length,
      node.start,
    );
    ctx.pushNumberingStyle({
      numId: numId,
      styleId: numberingStyle.styleId,
    });
    ctx.visitChildren(node);
    ctx.popNumberingStyle();
  },

  math: (node, ctx) => {
    let formula = normalizeFormula(node.value);
    let metrics = ctx.formulaMetrics.get(formulaKey(formula, true));

    let styleId = ctx.doc.styles.resource
      .getStyleByName("ispText_main")
      .getId();
    let paragraph = OXML.buildParagraphWithStyle(styleId);
    paragraph
      .getChild("w:pPr")
      .pushChild(XML.Node.build("w:jc").setAttr("w:val", "center"));

    if (!metrics) {
      throw new Error("Missing rendered display formula: " + formula);
    }
    let run = buildDualBlipImage(ctx, {
      pngBuffer: metrics.png,
      svgBuffer: metrics.svg,
      widthPt: Math.min(metrics.width, (MAX_BLOCK_WIDTH_CM * 72) / 2.54),
      heightPt:
        metrics.height *
        Math.min(1, (MAX_BLOCK_WIDTH_CM * 72) / 2.54 / metrics.width),
    });

    paragraph.pushChild(run);
    ctx.getCurrentNode().pushChild(paragraph);
  },

  inlineCode: (node, ctx) => {
    let styleId = ctx.doc.styles.resource
      .getStyleByName("ispListing Знак")
      .getId();

    let raw = OXML.buildRawTag(node.value, [
      XML.Node.build("w:rStyle").setAttr("w:val", styleId),
    ]);

    ctx.getCurrentNode().pushChild(raw);
  },

  html: (node, ctx) => {
    visitHtml(node, ctx);
  },
};
function vertAlignStyles(ctx: GenerationContext): XML.Node[] | undefined {
  let styles: XML.Node[] = [];
  if (ctx.boldDepth > 0) styles.push(XML.Node.build("w:b"));
  if (ctx.italicDepth > 0) styles.push(XML.Node.build("w:i"));
  let mode = ctx.vertAlignStack[ctx.vertAlignStack.length - 1];
  if (mode) styles.push(XML.Node.build("w:vertAlign").setAttr("w:val", mode));
  return styles.length ? styles : undefined;
}
function getTagClass(tag: string): string | null {
  let m =
    tag.match(/class\s*=\s*"([^"]*)"/) ||
    tag.match(/class\s*=\s*'([^']*)'/) ||
    tag.match(/class\s*=\s*([^\s>]+)/);
  return m ? m[1] : null;
}
function visitHtml(node: mdast.Html, ctx: GenerationContext) {
  let value = node.value.trim();
  let divMatch = value.match(/^<div\s+([^>]*)>([\s\S]*)<\/div>$/);
  if (divMatch) {
    let cls = getTagClass("<div " + divMatch[1] + ">");
    let inner = divMatch[2];
    if (cls === "img-caption") {
      emitCaption(ctx, inner, captionStyleId(ctx, "image"), false);
      return;
    }
    if (cls === "listing-caption" || cls === "table-caption") {
      emitCaption(ctx, inner, captionStyleId(ctx, "listing"), true);
      return;
    }
    emitCaption(
      ctx,
      inner,
      ctx.doc.styles.resource.getStyleByName("ispText_main").getId(),
      false,
    );
    return;
  }
  let openSpan = value.match(/^<span\s+([^>]*)>$/);
  if (openSpan) {
    let attrs = openSpan[1];
    let cls = getTagClass("<span " + attrs + ">");
    if (cls === "ref") {
      ctx.inlineRefMode = "ref";
      return;
    }
    if (cls === "cite") {
      ctx.inlineRefMode = "cite";
      return;
    }
    if (cls === "qed") {
      emitQed(ctx);
      return;
    }
    return;
  }
  if (/^<\/span>$/.test(value)) {
    ctx.inlineRefMode = null;
    return;
  }

  if (/^<sub>$/.test(value)) {
    ctx.vertAlignStack.push("subscript");
    return;
  }
  if (/^<sup>$/.test(value)) {
    ctx.vertAlignStack.push("superscript");
    return;
  }
  if (/^<\/sub>$/.test(value) || /^<\/sup>$/.test(value)) {
    ctx.vertAlignStack.pop();
    return;
  }
  if (/^<span\b[^>]*><\/span>$/.test(value)) {
    if (getTagClass(value) === "qed") emitQed(ctx);
    return;
  }
  if (/^<(em|i)>$/.test(value)) {
    ctx.italicDepth++;
    return;
  }
  if (/^<\/(em|i)>$/.test(value)) {
    ctx.italicDepth--;
    return;
  }
  if (/^<(strong|b)>$/.test(value)) {
    ctx.boldDepth++;
    return;
  }
  if (/^<\/(strong|b)>$/.test(value)) {
    ctx.boldDepth--;
    return;
  }
  if (/^<!--[\s\S]*-->$/.test(value)) return;
  throw new Error("Unsupported HTML: " + value);
}
function emitQed(ctx: GenerationContext) {
  let cur = ctx.getCurrentNode();
  let paragraph =
    cur.getTagName() === "w:p"
      ? cur
      : (() => {
          let styleId = ctx.doc.styles.resource
            .getStyleByName("ispText_main")
            .getId();
          let p = OXML.buildParagraphWithStyle(styleId);
          cur.pushChild(p);
          return p;
        })();
  paragraph.getChild("w:pPr").pushChild(OXML.buildQedTabStop());
  paragraph.pushChild(OXML.buildQedRun());
}
function captionStyleId(
  ctx: GenerationContext,
  kind: "image" | "listing",
): string {
  if (kind === "image") {
    return ctx.doc.styles.resource.getStyleByName("ispPicture_sign").getId();
  }
  return ctx.doc.styles.resource.getStyleByName("ispText_main").getId();
}
function emitCaption(
  ctx: GenerationContext,
  inner: string,
  styleId: string,
  italic: boolean,
) {
  let paragraph = OXML.buildParagraphWithStyle(styleId);
  if (italic) {
    paragraph
      .getChild("w:pPr")
      .pushChild(XML.Node.build("w:jc").setAttr("w:val", "left"));
  }
  let baseStyles = (): XML.Node[] =>
    italic
      ? [
          XML.Node.build("w:i"),
          XML.Node.build("w:iCs"),
          XML.Node.build("w:sz").setAttr("w:val", "18"),
          XML.Node.build("w:szCs").setAttr("w:val", "18"),
        ]
      : [];
  let vertAlign: ("subscript" | "superscript")[] = [];
  let emphasis = 0;
  let codeDepth = 0;
  let refMode: "ref" | "cite" | null = null;
  let pushRun = (text: string, mono: boolean) => {
    if (!text) return;
    let styles = baseStyles();
    let va = vertAlign[vertAlign.length - 1];
    if (va) styles.push(XML.Node.build("w:vertAlign").setAttr("w:val", va));
    if (emphasis > 0 && !italic) {
      styles.push(XML.Node.build("w:i"));
      styles.push(XML.Node.build("w:iCs"));
    }
    if (mono) {
      let styleId = ctx.doc.styles.resource
        .getStyleByName("ispListing Знак")
        .getId();
      styles.unshift(XML.Node.build("w:rStyle").setAttr("w:val", styleId));
    }
    paragraph.pushChild(
      OXML.buildRawTag(text, styles.length ? styles : undefined),
    );
  };
  let pushText = (raw: string) => {
    let text = decodeEntities(raw);
    if (!text) return;
    if (codeDepth > 0) {
      pushRun(text, true);
      return;
    }
    let parts = text.split("`");
    let balanced = parts.length % 2 === 1;
    parts.forEach((p, i) => pushRun(p, balanced && i % 2 === 1));
  };

  let tagRe = /<\/?[a-zA-Z][^>]*>/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = tagRe.exec(inner)) !== null) {
    let between = inner.slice(last, m.index);
    if (between) {
      if (refMode) {
        let id = decodeEntities(between).trim();
        let replacement =
          refMode === "cite"
            ? ctx.references.getCite(id)
            : ctx.references.getReference(id);
        paragraph.pushChild(
          OXML.buildRawTag(
            replacement,
            baseStyles().length ? baseStyles() : undefined,
          ),
        );
      } else {
        pushText(between);
      }
    }
    last = tagRe.lastIndex;

    let tag = m[0];
    let lower = tag.toLowerCase();
    if (/^<span\b/i.test(tag)) {
      let cls = getTagClass(tag);
      if (cls === "ref") refMode = "ref";
      else if (cls === "cite") refMode = "cite";
    } else if (lower === "</span>") {
      refMode = null;
    } else if (lower === "<sub>") vertAlign.push("subscript");
    else if (lower === "<sup>") vertAlign.push("superscript");
    else if (lower === "</sub>" || lower === "</sup>") vertAlign.pop();
    else if (lower === "<em>" || lower === "<i>") emphasis++;
    else if (lower === "</em>" || lower === "</i>")
      emphasis = Math.max(0, emphasis - 1);
    else if (lower === "<code>") codeDepth++;
    else if (lower === "</code>") codeDepth = Math.max(0, codeDepth - 1);
  }
  if (last < inner.length) pushText(inner.slice(last));

  ctx.getCurrentNode().pushChild(paragraph);
}

function decodeEntities(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function parseSizeAttr(attr: string | number | undefined) {
  if (typeof attr !== "string") return null;

  if (attr.endsWith("cm")) {
    let centimeters = parseFloat(attr.slice(0, -2));
    if (Number.isFinite(centimeters)) return centimeters;
  }

  return null;
}

function getImageSize(image: mdast.Image, aspect: number) {
  let width: number | null = parseSizeAttr(image.data?.attrs?.width);
  let height: number | null = parseSizeAttr(image.data?.attrs?.height);

  if (width === null && height === null) {
    width = 10;
  }

  if (height === null) {
    height = width / aspect;
  }

  if (width === null) {
    width = aspect * height;
  }

  return [width, height];
}

function createNumbering(
  document: WordDocument,
  abstractId: string,
  depth: number,
  start: number,
) {
  let newNumId = document.numbering.resource.getUnusedNumId();
  let newNumbering = new Num().readXml(XML.Node.build("w:num"));
  newNumbering.setAbstractNumId(abstractId);
  newNumbering.setId(newNumId);

  newNumbering
    .getLevelOverride(depth)
    .pushChild(
      XML.Node.build("w:startOverride").setAttr("w:val", String(start ?? 1)),
    );

  document.numbering.resource.nums.set(newNumId, newNumbering);
  return newNumId;
}
function addMediaRelationship(
  ctx: GenerationContext,
  name: string,
  data: Buffer,
): string {
  let mediaTarget = ctx.doc.getUniqueMediaTarget(name);
  let mediaPath = ctx.doc.getPathForTarget(mediaTarget);
  ctx.doc.saveFile(mediaPath, data);

  let id = ctx.doc.document.rels.getUnusedId();
  ctx.doc.document.rels.addRelation({
    id,
    target: mediaTarget,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  });
  return id;
}

export interface DualBlipOptions {
  pngBuffer: Buffer;
  svgBuffer?: Buffer | null;
  widthPt: number;
  heightPt: number;
  positionHalfPts?: number;
  pngName?: string;
}
function buildDualBlipImage(
  ctx: GenerationContext,
  opts: DualBlipOptions,
): XML.Node {
  const EMU = 914400 / 72;
  let widthEmu = Math.round(opts.widthPt * EMU);
  let heightEmu = Math.round(opts.heightPt * EMU);

  let baseName = opts.pngName ?? "img" + ctx.formulaCounter++;
  let pngBase = path
    .basename(baseName, path.extname(baseName))
    .replace(/[^a-zA-Z0-9._]/g, "_");

  let pngId = addMediaRelationship(ctx, pngBase + ".png", opts.pngBuffer);
  let blip = XML.Node.build("a:blip").setAttr("r:embed", pngId);

  if (opts.svgBuffer) {
    ctx.doc.contentTypes.resource.ensureDefault("svg", "image/svg+xml");
    let svgId = addMediaRelationship(ctx, pngBase + ".svg", opts.svgBuffer);

    blip.appendChildren([
      XML.Node.build("a:extLst").appendChildren([
        XML.Node.build("a:ext")
          .setAttr("uri", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
          .appendChildren([
            XML.Node.build("asvg:svgBlip")
              .setAttr(
                "xmlns:asvg",
                "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
              )
              .setAttr("r:embed", svgId),
          ]),
      ]),
    ]);
  }
  let numericId = ++ctx.formulaCounter;
  let drawing = createPicDrawing(numericId, widthEmu, heightEmu, blip);

  let run = XML.Node.build("w:r");
  if (opts.positionHalfPts !== undefined) {
    run.pushChild(
      XML.Node.build("w:rPr").appendChildren([
        XML.Node.build("w:position").setAttr(
          "w:val",
          String(opts.positionHalfPts),
        ),
      ]),
    );
  }
  run.pushChild(drawing);
  return run;
}

function createPicDrawing(
  id: number,
  widthEmu: number,
  heightEmu: number,
  blip: XML.Node,
) {
  return XML.Node.build("w:drawing").appendChildren([
    XML.Node.build("wp:inline")
      .setAttrs({
        distT: "0",
        distB: "0",
        distL: "0",
        distR: "0",
      })
      .appendChildren([
        XML.Node.build("wp:extent").setAttrs({
          cx: String(widthEmu),
          cy: String(heightEmu),
        }),
        XML.Node.build("wp:docPr").setAttrs({
          id: String(id),
          name: "Picture " + id,
        }),
        XML.Node.build("wp:cNvGraphicFramePr"),
        XML.Node.build("a:graphic").appendChildren([
          XML.Node.build("a:graphicData")
            .setAttrs({
              uri: "http://schemas.openxmlformats.org/drawingml/2006/picture",
            })
            .appendChildren([
              XML.Node.build("pic:pic").appendChildren([
                XML.Node.build("pic:nvPicPr").appendChildren([
                  XML.Node.build("pic:cNvPr").setAttrs({
                    id: String(id),
                    name: "Picture " + id,
                  }),
                  XML.Node.build("pic:cNvPicPr"),
                ]),
                XML.Node.build("pic:blipFill").appendChildren([
                  blip,
                  XML.Node.build("a:stretch").appendChildren([
                    XML.Node.build("a:fillRect"),
                  ]),
                ]),
                XML.Node.build("pic:spPr").appendChildren([
                  XML.Node.build("a:xfrm").appendChildren([
                    XML.Node.build("a:off").setAttrs({
                      x: "0",
                      y: "0",
                    }),
                    XML.Node.build("a:ext").setAttrs({
                      cx: String(widthEmu),
                      cy: String(heightEmu),
                    }),
                  ]),
                  XML.Node.build("a:prstGeom")
                    .setAttr("prst", "rect")
                    .appendChildren([XML.Node.build("a:avLst")]),
                ]),
              ]),
            ]),
        ]),
      ]),
  ]);
}
function readSvgSibling(url: string): Buffer | null {
  let svgPath = url.replace(/\.[^.]+$/, ".svg");
  if (svgPath !== url && fs.existsSync(svgPath)) {
    return fs.readFileSync(svgPath);
  }
  return null;
}
function visitBlockImage(node: mdast.Image, ctx: GenerationContext) {
  let pngPath =
    ctx.imagePaths.get(path.resolve(ctx.sourceDir, node.url)) ??
    path.resolve(ctx.sourceDir, node.url);
  let name = path.basename(pngPath, path.extname(pngPath));
  let svgPath = pngPath.replace(/\.[^.]+$/, ".svg");

  let styleId = ctx.doc.styles.resource.getStyleByName("ispText_main").getId();
  let paragraph = OXML.buildParagraphWithStyle(styleId);
  paragraph.getChild("w:pPr").appendChildren([
    XML.Node.build("w:keepNext"),
    XML.Node.build("w:keepLines"),
    XML.Node.build("w:spacing").setAttrs({
      "w:before": "60",
      "w:after": "60",
      "w:lineRule": "auto",
      "w:line": "240",
    }),
  ]);
  paragraph
    .getChild("w:pPr")
    .pushChild(XML.Node.build("w:jc").setAttr("w:val", "center"));

  if (!fs.existsSync(pngPath)) {
    throw new Error("Missing figure: " + pngPath);
  }

  let pngBuffer = fs.readFileSync(pngPath);
  let svgBuffer = fs.existsSync(svgPath) ? fs.readFileSync(svgPath) : null;

  let dim = ImageSize(pngBuffer);
  let aspect = dim.width / dim.height;
  let realWidthCm = (dim.width / FIGURE_DPI) * 2.54;
  let widthCm = Math.min(
    parseSizeAttr(node.data?.attrs?.width) ?? realWidthCm,
    MAX_BLOCK_WIDTH_CM,
  );
  let heightCm = widthCm / aspect;

  let run = buildDualBlipImage(ctx, {
    pngBuffer,
    svgBuffer,
    widthPt: (widthCm * 72) / 2.54,
    heightPt: (heightCm * 72) / 2.54,
    pngName: name + ".png",
  });

  paragraph.pushChild(run);
  ctx.getCurrentNode().pushChild(paragraph);
}

export function generateDocxBody(
  source: mdast.Root,
  target: WordDocument,
  meta: DocumentJsonMeta,
  references: DocumentReferences,
  formulaMetrics: Map<string, FormulaMetrics>,
  sourceDir: string,
  imagePaths = new Map<string, string>(),
) {
  new ParagraphTemplateSubstitution()
    .setDocument(target)
    .setTemplate("{{{body}}}")
    .setReplacement(() => {
      let context = new GenerationContext();
      let node = XML.Node.build("w:body");

      context.sourceDir = sourceDir;
      context.imagePaths = imagePaths;
      target.document.resource.toXml().visitSubtree("wp:docPr", (drawing) => {
        context.formulaCounter = Math.max(
          context.formulaCounter,
          Number(drawing.getAttr("id")) || 0,
        );
      });
      context.doc = target;
      context.meta = meta;
      context.references = references;
      if (formulaMetrics) context.formulaMetrics = formulaMetrics;
      context.pushNode(node);

      context.visitChildren(source);

      return node.getChildren();
    })
    .perform();
}

function getLinksParagraphs(doc: WordDocument, meta: DocumentJsonMeta) {
  let styleId = doc.styles.resource.getStyleByName("ispLitList").getId();
  let numId = "80";
  let linksSection = meta.getSection("links").asArray();

  let result = [];

  for (let link of linksSection) {
    let description: string;
    if (link.isMap()) {
      description = link.getString("description");
    } else {
      description = link.getString();
    }
    let paragraph = OXML.buildParagraphWithStyle(styleId);
    let style = paragraph.getChild("w:pPr");
    style.pushChild(OXML.buildNumPr("0", numId));

    paragraph.pushChild(OXML.buildRawTag(description));
    result.push(paragraph);
  }

  return result;
}

function getAuthors(
  doc: WordDocument,
  meta: DocumentJsonMeta,
  language: string,
) {
  let styleId = doc.styles.resource.getStyleByName("ispAuthor").getId();
  let authors = meta.getSection("authors").asArray();
  let organizations = meta
    .getSection("organizations")
    .asArray()
    .map((section) => section.getString("id"));

  let result = [];

  for (let author of authors) {
    let paragraph = OXML.buildParagraphWithStyle(styleId);

    let name = author.getString("name_" + language);
    let orcid = author.getString("orcid");
    let email = author.getString("email");
    let authorOrgs = author
      .getSection("organizations")
      .asArray()
      .map((section) => section.getString())
      .map((id) => organizations.indexOf(id) + 1)
      .join(",");

    const orcidPart = `, ORCID: ${orcid}`;
    let authorLine = `${name}${orcidPart}, <${email}>`;

    paragraph.pushChild(
      OXML.buildRawTag(authorOrgs + "\u00a0", [
        OXML.buildSuperscriptTextStyle(),
      ]),
    );
    paragraph.pushChild(OXML.buildRawTag(authorLine));

    result.push(paragraph);
  }

  return result;
}

function getOrganizations(
  doc: WordDocument,
  meta: DocumentJsonMeta,
  language: string,
) {
  let styleId = doc.styles.resource.getStyleByName("ispAuthor").getId();
  let organizations = meta.getSection("organizations").asArray();

  let orgIndex = 1;
  let result = [];

  for (let organization of organizations) {
    let paragraph = OXML.buildParagraphWithStyle(styleId);

    let indexLine = String(orgIndex);

    paragraph.pushChild(
      OXML.buildRawTag(indexLine, [OXML.buildSuperscriptTextStyle()]),
    );
    paragraph.pushChild(
      OXML.buildRawTag(organization.getString("name_" + language)),
    );

    result.push(paragraph);

    orgIndex++;
  }

  return result;
}

function getAuthorsDetail(doc: WordDocument, meta: DocumentJsonMeta) {
  let styleId = doc.styles.resource.getStyleByName("ispText_main").getId();
  let authors = meta.getSection("authors").asArray();

  let result = [];

  for (let author of authors) {
    for (let language of languages) {
      let line = author.getString("details_" + language);
      let newParagraph = OXML.buildParagraphWithStyle(styleId);
      newParagraph
        .getChild("w:pPr")
        .pushChild(
          XML.Node.build("w:spacing")
            .setAttr("w:before", "30")
            .setAttr("w:after", "120"),
        );
      newParagraph.pushChild(OXML.buildRawTag(line));
      result.push(newParagraph);
    }
  }

  return result;
}

export function substituteTemplates(
  document: WordDocument,
  meta: DocumentJsonMeta,
) {
  let inlineSubstitution = new InlineTemplateSubstitution().setDocument(
    document,
  );
  let paragraphSubstitution = new ParagraphTemplateSubstitution().setDocument(
    document,
  );

  for (let language of languages) {
    let templates = [
      "header",
      "abstract",
      "keywords",
      "for_citation",
      "acknowledgements",
    ];
    for (let template of templates) {
      let template_lang = template + "_" + language;
      let replacement = meta.getString(template_lang);

      inlineSubstitution
        .setTemplate("{{{" + template_lang + "}}}")
        .setReplacement(replacement)
        .perform();
    }

    let header = meta.getString("page_header_" + language);

    if (header === "@use_citation") {
      header = meta.getString("for_citation_" + language);
    }

    inlineSubstitution
      .setTemplate("{{{page_header_" + language + "}}}")
      .setReplacement(header)
      .perform();

    paragraphSubstitution
      .setTemplate("{{{authors_" + language + "}}}")
      .setReplacement(() => getAuthors(document, meta, language))
      .perform();

    paragraphSubstitution
      .setTemplate("{{{organizations_" + language + "}}}")
      .setReplacement(() => getOrganizations(document, meta, language))
      .perform();
  }

  paragraphSubstitution
    .setTemplate("{{{links}}}")
    .setReplacement(() => getLinksParagraphs(document, meta))
    .perform();

  paragraphSubstitution
    .setTemplate("{{{authors_detail}}}")
    .setReplacement(() => getAuthorsDetail(document, meta))
    .perform();
}
