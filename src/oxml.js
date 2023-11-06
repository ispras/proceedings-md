"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.buildParagraphTextTag = exports.buildSuperscriptTextStyle = exports.buildNumPr = exports.buildParagraphWithStyle = exports.getDocumentBody = void 0;
const XML = __importStar(require("./xml"));
function getDocumentBody(document) {
    return document.getChild("w:document").getChild("w:body");
}
exports.getDocumentBody = getDocumentBody;
function buildParagraphWithStyle(style) {
    return XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", style)
        ])
    ]);
}
exports.buildParagraphWithStyle = buildParagraphWithStyle;
function buildNumPr(ilvl, numId) {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>
    return XML.Node.build("w:numPr").appendChildren([
        XML.Node.build("w:ilvl").setAttr("w:val", "0"),
        XML.Node.build("w:numId").setAttr("w:val", numId),
    ]);
}
exports.buildNumPr = buildNumPr;
function buildSuperscriptTextStyle() {
    return XML.Node.build("w:vertAlign").setAttr("w:val", "superscript");
}
exports.buildSuperscriptTextStyle = buildSuperscriptTextStyle;
function buildParagraphTextTag(text, styles) {
    let result = XML.Node.build("w:r").appendChildren([
        XML.Node.build("w:t")
            .setAttr("xml:space", "preserve")
            .appendChildren([
            XML.Node.buildTextNode(text)
        ])
    ]);
    if (styles) {
        result.unshiftChild(XML.Node.build("w:rPr").appendChildren(styles));
    }
    return result;
}
exports.buildParagraphTextTag = buildParagraphTextTag;
//# sourceMappingURL=oxml.js.map