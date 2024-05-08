import * as XML from 'src/xml'
import {
    AnyElement,
    Block,
    Inline,
    MetaElement, MetaElementMap,
    metaElementToSource,
    MetaElementType,
    PandocJson,
    walkPandocElement
} from "src/pandoc/pandoc-json";

export function getOpenxmlInjection(node: XML.Node): MetaElement<"RawBlock"> {
    return {
        t: "RawBlock",
        c: ["openxml", node.toXmlString()]
    }
}

export default class PandocJsonPatcher {
    pandocJson: PandocJson

    constructor(pandocJson: PandocJson) {
        this.pandocJson = pandocJson
    }

    replaceDivWithClass(className: string, replacement: (contents: string) => Block) {
        this.replaceElements("Div", (element) => {
            if(element.c[0][1].indexOf(className) !== -1) {
                return replacement(metaElementToSource(element))
            }
        })
        return this
    }

    replaceSpanWithClass(className: string, replacement: (contents: string) => Inline) {
        this.replaceElements("Span", (element) => {
            if(element.c[0][1].indexOf(className) !== -1) {
                return replacement(metaElementToSource(element))
            }
        })
        return this
    }

    replaceElements<T extends MetaElementType>(type: T, callback: (element: MetaElement<T>) => AnyElement | void) {
        this.pandocJson.blocks = walkPandocElement(this.pandocJson.blocks, (element) => {
            if(element.t === type) {
                return callback(element as MetaElement<T>)
            }
        })
        return this
    }
}