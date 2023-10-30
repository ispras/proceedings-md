import {XMLBuilder, XMLParser} from "fast-xml-parser";

export const keys = {
    comment: "__comment__",
    text: "__text__",
    attributes: ":@"
}

export const parser = new XMLParser({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true,
    trimValues: false,
    commentPropName: keys.comment,
    textNodeName: keys.text
})

export const builder = new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    commentPropName: keys.comment,
    textNodeName: keys.text
})

export function buildAttributes(attributes: any): any {
    let result = {}
    result[keys.attributes] = attributes
    return result
}

export function buildTextNode(text: string): any {
    let result = {};
    result[keys.text] = text
    return result
}

export function getChildTag(tag: any, name: string): any {
    for (let child of tag) {
        if (child[name]) {
            return child
        }
    }
}

export function getTagName(tag: any): string {
    for (let key of Object.getOwnPropertyNames(tag)) {
        if (key === keys.attributes) continue
        return key
    }
}
