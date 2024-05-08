import path from "path";
import * as XML from "src/xml";

export const wordXmlns = new Map<string, string>([
    ["wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"],
    ["mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"],
    ["o", "urn:schemas-microsoft-com:office:office"],
    ["r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
    ["v", "urn:schemas-microsoft-com:vml"],
    ["wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"],
    ["wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
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
])

export const wordXmlnsIgnorable = new Set<string>(["wp14", "w14", "w15"])

export function getProperXmlnsForDocument(document: XML.Node) {
    let result = {}
    let ignorable = new Set<string>()

    document.visitSubtree((child) => {
        for(let name of XML.getUsedNames(child)) {
            let namespace = XML.getNamespace(name)

            if (!namespace || !wordXmlns.has(namespace)) {
                continue
            }

            result["xmlns:" + namespace] = wordXmlns.get(namespace)

            if (wordXmlnsIgnorable.has(namespace)) {
                ignorable.add(namespace)
            }
        }

        return true
    })

    if(ignorable.size) {
        result["xmlns:mc"] = wordXmlns.get("mc")
        result["mc:Ignorable"] = Array.from(ignorable).join(" ")
    }

    return result
}

export function buildParagraphWithStyle(style: string): XML.Node {
    return XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", style)
        ])
    ])
}

export function buildNumPr(ilvl: string, numId: string): XML.Node {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>

    return XML.Node.build("w:numPr").appendChildren([
        XML.Node.build("w:ilvl").setAttr("w:val", "0"),
        XML.Node.build("w:numId").setAttr("w:val", numId),
    ])
}

export function buildSuperscriptTextStyle(): XML.Node {
    return XML.Node.build("w:vertAlign").setAttr("w:val", "superscript")
}

export function buildParagraphTextTag(text: string, styles?: XML.Node[]): XML.Node {
    let result = XML.Node.build("w:r").appendChildren([
        XML.Node.build("w:t")
            .setAttr("xml:space", "preserve")
            .appendChildren([
                XML.Node.buildTextNode(text)
            ])
    ])

    if(styles) {
        result.unshiftChild(XML.Node.build("w:rPr").appendChildren(styles))
    }

    return result
}

export function getChildVal(node: XML.Node, tag: string) {
    let child = node.getChild(tag)
    if(child) return child.getAttr("w:val")
    return null
}

export function setChildVal(node: XML.Node, tag: string, value: string | null) {
    if(value === null) {
        node.removeChildren(tag)
    } else {
        let basedOnTag = node.getChild(tag)
        if(basedOnTag) basedOnTag.setAttr("w:val", value)
        else node.appendChildren([
            XML.Node.build(tag).setAttr("w:val", value)
        ])
    }
}

export function fixXmlns(document: XML.Node, rootTag: string) {
    document.getChild(rootTag).setAttrs(getProperXmlnsForDocument(document))
}

export function normalizePath(pathString: string) {
    pathString = path.posix.normalize(pathString)
    if(!pathString.startsWith("/")) {
        pathString = "/" + pathString
    }
    return pathString
}

export function getRelsPath(resourcePath: string) {
    let basename = path.basename(resourcePath)
    let dirname = path.dirname(resourcePath)

    return normalizePath(dirname + "/_rels/" + basename + ".rels")
}