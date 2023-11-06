"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Wrapper = exports.Serializable = exports.Node = exports.builder = exports.parser = exports.keys = void 0;
const fast_xml_parser_1 = require("fast-xml-parser");
exports.keys = {
    comment: "__comment__",
    text: "__text__",
    attributes: ":@",
    document: "__document__"
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
function checkFilter(filter, node) {
    if (!filter)
        return true;
    if (typeof filter === "string") {
        return node.getTagName() === filter;
    }
    return filter(node);
}
function getVisitArgs(args) {
    let filter = null;
    let callback = args[0];
    let startPosition = args[1];
    if (typeof args[1] === "function") {
        filter = args[0];
        callback = args[1];
        startPosition = args[2];
    }
    return {
        filter: filter,
        callback: callback,
        startPosition: startPosition
    };
}
class Node {
    element;
    tempDestroyed = false;
    constructor(element) {
        if (Array.isArray(element)) {
            throw new Error("XML.Node must be constructed from the xml object, not its children list");
        }
        this.element = element;
    }
    getTagName() {
        this.checkTemporary();
        for (let key of Object.getOwnPropertyNames(this.element)) {
            // Be safe from prototype pollution
            if (key === "__proto__" || key === exports.keys.attributes)
                continue;
            return key;
        }
        return null;
    }
    pushChild(child) {
        this.checkTemporary();
        let tagName = this.getTagName();
        if (!Array.isArray(this.element[tagName])) {
            throw new Error("Cannot call pushChild on " + tagName + " element");
        }
        let childArray = this.element[tagName];
        childArray.push(child.raw());
    }
    unshiftChild(child) {
        this.checkTemporary();
        let tagName = this.getTagName();
        if (!Array.isArray(this.element[tagName])) {
            throw new Error("Cannot call pushChild on " + tagName + " element");
        }
        let childArray = this.element[tagName];
        childArray.unshift(child.raw());
    }
    getChildren(filter = null) {
        this.checkTemporary();
        let result = [];
        this.visitChildren(filter, (child) => {
            result.push(child.shallowCopy());
        });
        return result;
    }
    getChild(arg = null) {
        this.checkTemporary();
        if (Array.isArray(arg)) {
            let path = arg;
            if (path.length === 0) {
                return this;
            }
            let result = new Node(this.element);
            for (let i = 0; i < path.length; i++) {
                if (!result.element)
                    return null;
                let tagName = result.getTagName();
                let pathComponent = path[i];
                let children = result.element[tagName];
                if (pathComponent < 0) {
                    result.element = children[children.length + pathComponent];
                }
                else {
                    result.element = children[pathComponent];
                }
            }
            if (!result.element)
                return null;
            return result;
        }
        else {
            let filter = arg;
            let result = null;
            this.visitChildren(filter, (child) => {
                if (result) {
                    throw new Error("Element have multiple children matching the given filter");
                }
                result = child.shallowCopy();
            });
            return result;
        }
    }
    visitChildren(...args) {
        this.checkTemporary();
        let { filter, callback, startPosition } = getVisitArgs(args);
        let tagName = this.getTagName();
        if (!Array.isArray(this.element[tagName])) {
            return;
        }
        let index = startPosition ?? 0;
        let tmpNode = new Node(null);
        for (let child of this.element[tagName]) {
            tmpNode.element = child;
            if (checkFilter(filter, tmpNode)) {
                if (callback(tmpNode, index) === false) {
                    break;
                }
            }
            index++;
        }
        tmpNode.markDestroyed();
    }
    visitSubtree(...args) {
        this.checkTemporary();
        let { filter, callback, startPosition } = getVisitArgs(args);
        let tmpNode = new Node(null);
        let startPath = startPosition ?? [];
        let startDepth = 0;
        let path = [];
        const walk = (node) => {
            let tagName = node.getTagName();
            let children = node.element[tagName];
            if (!Array.isArray(children)) {
                return;
            }
            let depth = path.length;
            let startIndex = 0;
            if (depth < startDepth && startPath.length) {
                startIndex = startPath[startPath.length];
                startDepth = depth;
            }
            for (let index = startIndex; index < children.length; index++) {
                path.push(index);
                tmpNode.element = children[index];
                let filterPass = checkFilter(filter, tmpNode);
                let goDeeper = true;
                if (filterPass) {
                    goDeeper = callback(tmpNode, path) === true;
                }
                if (goDeeper) {
                    walk(tmpNode);
                }
                path.pop();
            }
        };
        walk(this);
        tmpNode.markDestroyed();
    }
    removeChild(path) {
        let topIndex = path.pop();
        let child = this.getChild(path);
        let tagName = child.getTagName();
        child.element[tagName].splice(topIndex, 1);
        path.push(topIndex);
    }
    removeChildren(filter = null) {
        this.checkTemporary();
        let tagName = this.getTagName();
        if (!Array.isArray(this.element[tagName])) {
            return;
        }
        let node = new Node(null);
        for (let i = 0; i < this.element[tagName].length; i++) {
            node.element = this.element[tagName][i];
            if (checkFilter(filter, node)) {
                this.element[tagName].splice(i, 1);
                i--;
            }
        }
        node.markDestroyed();
    }
    isTextNode() {
        this.checkTemporary();
        return this.getTagName() == exports.keys.text;
    }
    isCommentNode() {
        this.checkTemporary();
        return this.getTagName() == exports.keys.comment;
    }
    getText() {
        this.checkTemporary();
        if (!this.isTextNode()) {
            throw new Error("getText() is called on " + this.getTagName() + " element");
        }
        return String(this.element[exports.keys.text]);
    }
    setText(text) {
        this.checkTemporary();
        if (!this.isTextNode()) {
            throw new Error("setText() is called on " + this.getTagName() + " element");
        }
        this.element[exports.keys.text] = text;
    }
    getComment() {
        this.checkTemporary();
        if (!this.isCommentNode()) {
            throw new Error("getComment() is called on " + this.getTagName() + " element");
        }
        let textChild = this.getChild(exports.keys.text);
        return textChild.getText();
    }
    static build(tagName) {
        let element = {};
        element[tagName] = [];
        return new Node(element);
    }
    static createDocument(args = {}) {
        args = Object.assign({
            version: "1.0",
            encoding: "UTF-8",
            standalone: "yes"
        }, args);
        let document = this.build(exports.keys.document);
        document.appendChildren([
            Node.build("?xml")
                .setAttrs(args)
                .appendChildren([
                Node.buildTextNode("")
            ])
        ]);
        return document;
    }
    static buildTextNode(text) {
        let element = {};
        element[exports.keys.text] = text;
        return new Node(element);
    }
    setAttr(attribute, value) {
        this.checkTemporary();
        if (!this.element[exports.keys.attributes]) {
            this.element[exports.keys.attributes] = {};
        }
        this.element[exports.keys.attributes][attribute] = value;
        return this;
    }
    setAttrs(attributes) {
        this.checkTemporary();
        this.element[exports.keys.attributes] = attributes;
        return this;
    }
    getAttr(attribute) {
        this.checkTemporary();
        if (!this.element[exports.keys.attributes]) {
            return undefined;
        }
        return String(this.element[exports.keys.attributes][attribute]);
    }
    clearChildren(path = []) {
        this.checkTemporary();
        let parent = this.getChild(path);
        parent.element[parent.getTagName()] = [];
        return this;
    }
    insertChildren(children, path) {
        this.checkTemporary();
        let insertIndex = path.pop();
        let parent = this.getChild(path);
        path.push(insertIndex);
        let lastChildren = parent.element[parent.getTagName()];
        if (insertIndex < 0) {
            insertIndex = children.length + insertIndex + 1;
        }
        lastChildren.splice(insertIndex, 0, ...children.map(child => child.raw()));
        return this;
    }
    appendChildren(children, path = []) {
        path.push(-1);
        this.insertChildren(children, path);
        path.pop();
        return this;
    }
    unshiftChildren(children, path = []) {
        path.push(0);
        this.insertChildren(children, path);
        path.pop();
        return this;
    }
    assign(another) {
        this.checkTemporary();
        if (this === another) {
            return;
        }
        if (this.element) {
            this.element[this.getTagName()] = undefined;
        }
        else {
            this.element = {};
        }
        this.element[another.getTagName()] = JSON.parse(JSON.stringify(another.raw()[another.getTagName()]));
        if (another.raw()[exports.keys.attributes]) {
            this.element[exports.keys.attributes] = JSON.parse(JSON.stringify(another.raw()[exports.keys.attributes]));
        }
        else {
            this.element[exports.keys.attributes] = {};
        }
        return this;
    }
    static fromXmlString(str) {
        let object = exports.parser.parse(str);
        let wrapped = {};
        wrapped[exports.keys.document] = object;
        return new Node(wrapped);
    }
    toXmlString() {
        this.checkTemporary();
        if (this.getTagName() === exports.keys.document) {
            return exports.builder.build(this.element[exports.keys.document]);
        }
        else {
            return exports.builder.build([this.element]);
        }
    }
    raw() {
        this.checkTemporary();
        return this.element;
    }
    checkTemporary() {
        if (this.tempDestroyed) {
            throw new Error("Method access to an outdated temporary Node. Make sure to call .shallowCopy() on temporary " +
                "nodes before accessing them outside your visitChildren/visitSubtree body scope");
        }
    }
    markDestroyed() {
        // From now on, the checkTemporary method will throw
        this.tempDestroyed = true;
    }
    shallowCopy() {
        this.checkTemporary();
        return new Node(this.element);
    }
    deepCopy() {
        return new Node(null).assign(this);
    }
}
exports.Node = Node;
class Serializable {
    readXmlString(xmlString) {
        this.readXml(Node.fromXmlString(xmlString));
        return this;
    }
    readXml(xml) {
        throw new Error("readXml is not implemented");
        return this;
    }
    toXmlString() {
        return this.toXml().toXmlString();
    }
    toXml() {
        throw new Error("toXml is not implemented");
    }
}
exports.Serializable = Serializable;
class Wrapper extends Serializable {
    node = null;
    readXml(xml) {
        this.node = xml;
        return this;
    }
    toXml() {
        return this.node;
    }
}
exports.Wrapper = Wrapper;
//# sourceMappingURL=xml.js.map