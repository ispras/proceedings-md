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
exports.buildParagraphTextTag = exports.buildSuperscriptTextStyle = exports.buildNumPr = exports.buildParagraphWithStyle = void 0;
const XML = __importStar(require("./xml"));
function buildParagraphWithStyle(style) {
    return {
        "w:p": [{
                "w:pPr": [{
                        "w:pStyle": [],
                        ...XML.buildAttributes({ "w:val": style })
                    }]
            }]
    };
}
exports.buildParagraphWithStyle = buildParagraphWithStyle;
function buildNumPr(ilvl, numId) {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>
    return {
        "w:numPr": [{
                "w:ilvl": [],
                ...XML.buildAttributes({ "w:val": "0" })
            }, {
                "w:numId": [],
                ...XML.buildAttributes({ "w:val": numId })
            }]
    };
}
exports.buildNumPr = buildNumPr;
function buildSuperscriptTextStyle() {
    return {
        "w:vertAlign": [],
        ...XML.buildAttributes({ "w:val": "superscript" })
    };
}
exports.buildSuperscriptTextStyle = buildSuperscriptTextStyle;
function buildParagraphTextTag(text, styles) {
    let result = {
        "w:r": [
            {
                "w:t": [XML.buildTextNode(text)],
                ...XML.buildAttributes({ "xml:space": "preserve" })
            }
        ]
    };
    if (styles) {
        result["w:r"].unshift({
            "w:rPr": styles
        });
    }
    return result;
}
exports.buildParagraphTextTag = buildParagraphTextTag;
//# sourceMappingURL=oxml.js.map