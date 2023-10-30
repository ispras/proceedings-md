"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getTagName = exports.getChildTag = exports.buildTextNode = exports.buildAttributes = exports.builder = exports.parser = exports.keys = void 0;
const fast_xml_parser_1 = require("fast-xml-parser");
exports.keys = {
    comment: "__comment__",
    text: "__text__",
    attributes: ":@"
};
exports.parser = new fast_xml_parser_1.XMLParser({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true,
    trimValues: false,
    commentPropName: exports.keys.comment,
    textNodeName: exports.keys.text
});
exports.builder = new fast_xml_parser_1.XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    commentPropName: exports.keys.comment,
    textNodeName: exports.keys.text
});
function buildAttributes(attributes) {
    let result = {};
    result[exports.keys.attributes] = attributes;
    return result;
}
exports.buildAttributes = buildAttributes;
function buildTextNode(text) {
    let result = {};
    result[exports.keys.text] = text;
    return result;
}
exports.buildTextNode = buildTextNode;
function getChildTag(tag, name) {
    for (let child of tag) {
        if (child[name]) {
            return child;
        }
    }
}
exports.getChildTag = getChildTag;
function getTagName(tag) {
    for (let key of Object.getOwnPropertyNames(tag)) {
        if (key === exports.keys.attributes)
            continue;
        return key;
    }
}
exports.getTagName = getTagName;
//# sourceMappingURL=xml.js.map