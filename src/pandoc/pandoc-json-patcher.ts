import * as XML from 'src/xml'
import {MetaElement, metaElementToSource, PandocJson, walkPandocElement} from "src/pandoc/pandoc-json";

function getOpenxmlInjection(xml: string): MetaElement<"RawBlock"> {
    return {
        t: "RawBlock",
        c: ["openxml", xml]
    }
}

export default class PandocJsonPatcher {
    pandocJson: PandocJson

    constructor(pandocJson: PandocJson) {
        this.pandocJson = pandocJson
    }

    replaceDivWithClass(className: string, replacement: (contents: string) => XML.Node) {
        this.pandocJson.blocks = walkPandocElement(this.pandocJson.blocks, (element) => {
            if(element.t === "Div" && element.c[0][1].indexOf(className) !== -1) {
                return getOpenxmlInjection(replacement(metaElementToSource(element)).toXmlString())
            }
        })
        return this
    }
}