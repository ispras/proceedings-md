import path from "path";
import * as XML from "../xml";

export const wordXmlns = new Map<string, string>([
  ["wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"],
  ["mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"],
  ["o", "urn:schemas-microsoft-com:office:office"],
  ["r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
  ["m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
  ["v", "urn:schemas-microsoft-com:vml"],
  [
    "wp14",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  ],
  [
    "wp",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  ],
  ["w10", "urn:schemas-microsoft-com:office:word"],
  ["w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
  ["w14", "http://schemas.microsoft.com/office/word/2010/wordml"],
  ["w15", "http://schemas.microsoft.com/office/word/2012/wordml"],
  ["wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"],
  ["wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"],
  ["wne", "http://schemas.microsoft.com/office/word/2006/wordml"],
  ["wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"],
  ["pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
  ["a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
]);

export const wordXmlnsIgnorable = new Set<string>(["wp14", "w14", "w15"]);

export function getProperXmlnsForDocument(document: XML.Node) {
  let result = {};
  let ignorable = new Set<string>();

  document.visitSubtree((child) => {
    for (let name of XML.getUsedNames(child)) {
      let namespace = XML.getNamespace(name);

      if (!namespace || !wordXmlns.has(namespace)) {
        continue;
      }

      result["xmlns:" + namespace] = wordXmlns.get(namespace);

      if (wordXmlnsIgnorable.has(namespace)) {
        ignorable.add(namespace);
      }
    }

    return true;
  });

  if (ignorable.size) {
    result["xmlns:mc"] = wordXmlns.get("mc");
    result["mc:Ignorable"] = Array.from(ignorable).join(" ");
  }

  return result;
}

export function buildParagraphWithStyle(style: string): XML.Node {
  return XML.Node.build("w:p").appendChildren([
    XML.Node.build("w:pPr").appendChildren([
      XML.Node.build("w:pStyle").setAttr("w:val", style),
    ]),
  ]);
}

export function buildNumPr(ilvl: string, numId: string): XML.Node {
  return XML.Node.build("w:numPr").appendChildren([
    XML.Node.build("w:ilvl").setAttr("w:val", ilvl),
    XML.Node.build("w:numId").setAttr("w:val", numId),
  ]);
}

export function buildSuperscriptTextStyle(): XML.Node {
  return XML.Node.build("w:vertAlign").setAttr("w:val", "superscript");
}

export function buildLineBreak() {
  return XML.Node.build("w:br");
}

export function buildTextTag(text: string): XML.Node {
  return XML.Node.build("w:t")
    .setAttr("xml:space", "preserve")
    .appendChildren([XML.Node.buildTextNode(text)]);
}

export function buildRawTag(text: string, styles?: XML.Node[]): XML.Node {
  let result = XML.Node.build("w:r");

  if (styles) {
    result.pushChild(XML.Node.build("w:rPr").appendChildren(styles));
  }

  result.pushChild(buildTextTag(text));

  return result;
}
const TEXT_AREA_WIDTH_TWIPS = 7938;
export function buildQedTabStop(): XML.Node {
  return XML.Node.build("w:tabs").appendChildren([
    XML.Node.build("w:tab")
      .setAttr("w:val", "right")
      .setAttr("w:pos", String(TEXT_AREA_WIDTH_TWIPS)),
  ]);
}
export function buildQedRun(): XML.Node {
  return XML.Node.build("w:r").appendChildren([
    XML.Node.build("w:tab"),
    buildTextTag("□"),
  ]);
}

export function getChildVal(node: XML.Node, tag: string) {
  let child = node.getChild(tag);
  if (child) return child.getAttr("w:val");
  return null;
}

export function setChildVal(node: XML.Node, tag: string, value: string | null) {
  if (value === null) {
    node.removeChildren(tag);
  } else {
    let basedOnTag = node.getChild(tag);
    if (basedOnTag) basedOnTag.setAttr("w:val", value);
    else node.appendChildren([XML.Node.build(tag).setAttr("w:val", value)]);
  }
}

export function fixXmlns(document: XML.Node, rootTag: string) {
  const root = document.getChild(rootTag);
  root.setAttrs({ ...root.getAttrs(), ...getProperXmlnsForDocument(document) });
}

export function normalizePath(pathString: string) {
  pathString = path.posix.normalize(pathString);
  if (!pathString.startsWith("/")) {
    pathString = "/" + pathString;
  }
  return pathString;
}

export function getRelsPath(resourcePath: string) {
  let basename = path.basename(resourcePath);
  let dirname = path.dirname(resourcePath);

  return normalizePath(dirname + "/_rels/" + basename + ".rels");
}
