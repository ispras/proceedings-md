import {XMLBuilder, XMLParser} from "fast-xml-parser";
import * as XML from "./xml";

export function buildParagraphWithStyle(style: string): any {
    return {
        "w:p": [{
            "w:pPr": [{
                "w:pStyle": [],
                ...XML.buildAttributes({"w:val": style})
            }]
        }]
    };
}

export function buildNumPr(ilvl: string, numId: string): any {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>

    return {
        "w:numPr": [{
            "w:ilvl": [],
            ...XML.buildAttributes({"w:val": "0"})
        }, {
            "w:numId": [],
            ...XML.buildAttributes({"w:val": numId})
        }]
    }
}

export function buildSuperscriptTextStyle(): any {
    return {
        "w:vertAlign": [],
        ...XML.buildAttributes({"w:val": "superscript"})
    }
}

export function buildParagraphTextTag(text: string, styles?: any[]): any {
    let result = {
        "w:r": [
            {
                "w:t": [XML.buildTextNode(text)],
                ...XML.buildAttributes({"xml:space": "preserve"})
            }
        ]
    }

    if(styles) {
        result["w:r"].unshift({
            "w:rPr": styles
        })
    }

    return result;
}