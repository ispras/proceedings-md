
import * as XML from "./xml";

export function getDocumentBody(document: XML.Node): XML.Node {
    return document.getChild("w:document").getChild("w:body")
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