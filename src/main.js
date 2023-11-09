'use strict';

var path = require('path');
var fs = require('fs');
var child_process = require('child_process');
var fastXmlParser = require('fast-xml-parser');
var process = require('process');
var JSZip = require('jszip');

function _interopNamespaceDefault(e) {
    var n = Object.create(null);
    if (e) {
        Object.keys(e).forEach(function (k) {
            if (k !== 'default') {
                var d = Object.getOwnPropertyDescriptor(e, k);
                Object.defineProperty(n, k, d.get ? d : {
                    enumerable: true,
                    get: function () { return e[k]; }
                });
            }
        });
    }
    n.default = e;
    return Object.freeze(n);
}

var path__namespace = /*#__PURE__*/_interopNamespaceDefault(path);
var fs__namespace = /*#__PURE__*/_interopNamespaceDefault(fs);
var process__namespace = /*#__PURE__*/_interopNamespaceDefault(process);

function pandoc(src, args) {
    return new Promise((resolve, reject) => {
        let stdout = "";
        let stderr = "";
        let pandocProcess = child_process.spawn('pandoc', args);
        pandocProcess.stdin.end(src, 'utf-8');
        pandocProcess.stdout.on('data', (data) => {
            stdout += data;
        });
        pandocProcess.stderr.on('data', (data) => {
            stderr += data;
        });
        pandocProcess.on('exit', function (code) {
            if (stderr.length) {
                console.error("There was some pandoc warnings along the way:");
                console.error(stderr);
            }
            if (code == 0) {
                resolve(stdout);
            }
            else {
                reject(new Error("Pandoc returned non-zero exit code"));
            }
        });
    });
}
async function markdownToPandocJson(markdown, flags) {
    let meta = await pandoc(markdown, ["-f", "markdown", "-t", "json", ...flags]);
    return JSON.parse(meta);
}
async function pandocJsonToDocx(pandocJson, flags) {
    return await pandoc(JSON.stringify(pandocJson), ["-f", "json", "-t", "docx", ...flags]);
}
const tokenClasses = [
    "KeywordTok",
    "NormalTok",
    "OperatorTok",
    "DataTypeTok",
    "PreprocessorTok",
    "DecValTok",
    "BaseNTok",
    "FloatTok",
    "ConstantTok",
    "CharTok",
    "SpecialCharTok",
    "StringTok",
    "VerbatimStringTok",
    "SpecialStringTok",
    "ImportTok",
    "CommentTok",
    "DocumentationTok",
    "AnnotationTok",
    "CommentVarTok",
    "OtherTok",
    "FunctionTok",
    "VariableTok",
    "ControlFlowTok",
    "BuiltInTok",
    "ExtensionTok",
    "AttributeTok",
    "RegionMarkerTok",
    "InformationTok",
    "WarningTok",
    "AlertTok",
    "ErrorTok"
];

const keys = {
    comment: "__comment__",
    text: "__text__",
    attributes: ":@",
    document: "__document__"
};
const parser = new fastXmlParser.XMLParser({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true,
    trimValues: false,
    commentPropName: keys.comment,
    textNodeName: keys.text
});
const builder = new fastXmlParser.XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    commentPropName: keys.comment,
    textNodeName: keys.text
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
            if (key === "__proto__" || key === keys.attributes)
                continue;
            return key;
        }
        return null;
    }
    pushChild(child) {
        this.checkTemporary();
        let children = this.getRawChildren();
        if (children === null) {
            throw new Error("Cannot call pushChild on " + this.getTagName() + " element");
        }
        children.push(child.raw());
    }
    unshiftChild(child) {
        this.checkTemporary();
        let children = this.getRawChildren();
        if (children === null) {
            throw new Error("Cannot call unshiftChild on " + this.getTagName() + " element");
        }
        children.unshift(child.raw());
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
        if (path.length === 0) {
            throw new Error("Cannot call removeChild with empty path");
        }
        let topIndex = path.pop();
        let child = this.getChild(path);
        let childChildren = child.getRawChildren();
        if (childChildren === null) {
            throw new Error("Cannot call removeChild for " + child.getTagName() + " element");
        }
        childChildren.splice(topIndex, 1);
        path.push(topIndex);
    }
    removeChildren(filter = null) {
        this.checkTemporary();
        let children = this.getRawChildren();
        if (children === null) {
            throw new Error("Cannot call removeChildren on " + this.getTagName() + " element");
        }
        let node = new Node(null);
        for (let i = 0; i < children.length; i++) {
            node.element = children[i];
            if (checkFilter(filter, node)) {
                children.splice(i, 1);
                i--;
            }
        }
        node.markDestroyed();
    }
    isTextNode() {
        this.checkTemporary();
        return this.getTagName() == keys.text;
    }
    isCommentNode() {
        this.checkTemporary();
        return this.getTagName() == keys.comment;
    }
    getText() {
        this.checkTemporary();
        if (!this.isTextNode()) {
            throw new Error("getText() is called on " + this.getTagName() + " element");
        }
        return String(this.element[keys.text]);
    }
    setText(text) {
        this.checkTemporary();
        if (!this.isTextNode()) {
            throw new Error("setText() is called on " + this.getTagName() + " element");
        }
        this.element[keys.text] = text;
    }
    getComment() {
        this.checkTemporary();
        if (!this.isCommentNode()) {
            throw new Error("getComment() is called on " + this.getTagName() + " element");
        }
        let textChild = this.getChild(keys.text);
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
        let document = this.build(keys.document);
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
        element[keys.text] = text;
        return new Node(element);
    }
    setAttr(attribute, value) {
        this.checkTemporary();
        if (!this.element[keys.attributes]) {
            this.element[keys.attributes] = {};
        }
        this.element[keys.attributes][attribute] = value;
        return this;
    }
    setAttrs(attributes) {
        this.checkTemporary();
        this.element[keys.attributes] = attributes;
        return this;
    }
    getAttrs() {
        if (!this.element[keys.attributes]) {
            this.element[keys.attributes] = {};
        }
        return this.element[keys.attributes];
    }
    getAttr(attribute) {
        this.checkTemporary();
        let attrs = this.getAttrs();
        let attr = attrs[attribute];
        if (attr === undefined)
            return null;
        return String(attr);
    }
    clearChildren(path = []) {
        this.checkTemporary();
        let parent = this.getChild(path);
        parent.element[parent.getTagName()] = [];
        return this;
    }
    insertChildren(children, path) {
        this.checkTemporary();
        if (path.length === 0) {
            throw new Error("Cannot call insertChildren with empty path");
        }
        let insertIndex = path.pop();
        let parent = this.getChild(path);
        path.push(insertIndex);
        let lastChildren = parent.getRawChildren();
        if (lastChildren === null) {
            throw new Error("Cannot call insertChildren for " + parent.getTagName() + " element");
        }
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
        if (another.raw()[keys.attributes]) {
            this.element[keys.attributes] = JSON.parse(JSON.stringify(another.raw()[keys.attributes]));
        }
        else {
            this.element[keys.attributes] = {};
        }
        return this;
    }
    static fromXmlString(str) {
        let object = parser.parse(str);
        let wrapped = {};
        wrapped[keys.document] = object;
        return new Node(wrapped);
    }
    toXmlString() {
        this.checkTemporary();
        if (this.getTagName() === keys.document) {
            return builder.build(this.element[keys.document]);
        }
        else {
            return builder.build([this.element]);
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
        this.checkTemporary();
        // From now on, the checkTemporary method will throw
        this.tempDestroyed = true;
    }
    getRawContents() {
        this.checkTemporary();
        return this.element[this.getTagName()];
    }
    getRawChildren() {
        this.checkTemporary();
        let contents = this.getRawContents();
        if (Array.isArray(contents)) {
            return contents;
        }
        return null;
    }
    shallowCopy() {
        this.checkTemporary();
        return new Node(this.element);
    }
    deepCopy() {
        this.checkTemporary();
        return new Node(null).assign(this);
    }
    isLeaf() {
        this.checkTemporary();
        return this.getRawChildren() === null;
    }
    getChildrenCount() {
        return this.getRawChildren()?.length ?? 0;
    }
}
class Serializable {
    readXmlString(xmlString) {
        this.readXml(Node.fromXmlString(xmlString));
        return this;
    }
    readXml(xml) {
        throw new Error("readXml is not implemented");
    }
    toXmlString() {
        return this.toXml().toXmlString();
    }
    toXml() {
        throw new Error("toXml is not implemented");
    }
}
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
function getNamespace(name) {
    let parts = name.split(":");
    if (parts.length >= 2) {
        return parts[0];
    }
    return null;
}
function* getUsedNames(tag) {
    let tagName = tag.getTagName();
    yield tagName;
    let attributes = tag.getAttrs();
    for (let key of Object.getOwnPropertyNames(attributes)) {
        // Be safe from prototype pollution
        if (key === "__proto__")
            continue;
        yield key;
    }
}
function getTextContents(tag) {
    let result = "";
    tag.visitSubtree(keys.text, (node) => {
        result += node.getText();
    });
    return result;
}

const wordXmlns = new Map([
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
]);
const wordXmlnsIgnorable = new Set(["wp14", "w14", "w15"]);
function getProperXmlnsForDocument(document) {
    let result = {};
    let ignorable = new Set();
    document.visitSubtree((child) => {
        for (let name of getUsedNames(child)) {
            let namespace = getNamespace(name);
            if (!namespace || !wordXmlns.has(namespace)) {
                continue;
            }
            result["xmlns:" + namespace] = wordXmlns.get(namespace);
            if (wordXmlnsIgnorable.has(namespace)) {
                ignorable.add(namespace);
            }
        }
        return true;
    });
    if (ignorable.size) {
        result["xmlns:mc"] = wordXmlns.get("mc");
        result["mc:Ignorable"] = Array.from(ignorable).join(" ");
    }
    return result;
}
function buildParagraphWithStyle(style) {
    return Node.build("w:p").appendChildren([
        Node.build("w:pPr").appendChildren([
            Node.build("w:pStyle").setAttr("w:val", style)
        ])
    ]);
}
function buildNumPr(ilvl, numId) {
    // <w:numPr>
    //    <w:ilvl w:val="<ilvl>"/>
    //    <w:numId w:val="<numId>"/>
    // </w:numPr>
    return Node.build("w:numPr").appendChildren([
        Node.build("w:ilvl").setAttr("w:val", "0"),
        Node.build("w:numId").setAttr("w:val", numId),
    ]);
}
function buildSuperscriptTextStyle() {
    return Node.build("w:vertAlign").setAttr("w:val", "superscript");
}
function buildParagraphTextTag(text, styles) {
    let result = Node.build("w:r").appendChildren([
        Node.build("w:t")
            .setAttr("xml:space", "preserve")
            .appendChildren([
            Node.buildTextNode(text)
        ])
    ]);
    if (styles) {
        result.unshiftChild(Node.build("w:rPr").appendChildren(styles));
    }
    return result;
}
function getChildVal(node, tag) {
    let child = node.getChild(tag);
    if (child)
        return child.getAttr("w:val");
    return null;
}
function setChildVal(node, tag, value) {
    if (value === null) {
        node.removeChildren(tag);
    }
    else {
        let basedOnTag = node.getChild(tag);
        if (basedOnTag)
            basedOnTag.setAttr("w:val", value);
        else
            node.appendChildren([
                Node.build(tag).setAttr("w:val", value)
            ]);
    }
}
function fixXmlns(document, rootTag) {
    document.getChild(rootTag).setAttrs(getProperXmlnsForDocument(document));
}
function normalizePath(pathString) {
    pathString = path.posix.normalize(pathString);
    if (!pathString.startsWith("/")) {
        pathString = "/" + pathString;
    }
    return pathString;
}
function getRelsPath(resourcePath) {
    let basename = path.basename(resourcePath);
    let dirname = path.dirname(resourcePath);
    return normalizePath(dirname + "/_rels/" + basename + ".rels");
}

class Relationships extends Serializable {
    relations = new Map();
    readXml(xml) {
        this.relations = new Map();
        xml.getChild("Relationships")?.visitChildren("Relationship", (child) => {
            let id = child.getAttr("Id");
            let type = child.getAttr("Type");
            let target = child.getAttr("Target");
            if (id !== undefined && type !== undefined && target !== undefined) {
                this.relations.set(id, {
                    id: id,
                    type: type,
                    target: target
                });
            }
        });
        return this;
    }
    toXml() {
        let relations = Array.from(this.relations.values());
        return Node.createDocument().appendChildren([
            Node.build("Relationships")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
                .appendChildren(relations.map((relation) => {
                return Node.build("Relationship")
                    .setAttr("Id", relation.id)
                    .setAttr("Type", relation.type)
                    .setAttr("Target", relation.target);
            }))
        ]);
    }
    getUnusedId() {
        let prefix = "rId";
        let index = 1;
        while (this.relations.has(prefix + index)) {
            index++;
        }
        return prefix + index;
    }
    getRelForTarget(target) {
        for (let rel of this.relations.values()) {
            if (rel.target === target) {
                return rel;
            }
        }
    }
}

class ContentTypes extends Serializable {
    defaults;
    overrides;
    readXml(xml) {
        this.defaults = [];
        this.overrides = [];
        let types = xml.getChild("Types");
        types?.visitChildren("Default", (child) => {
            let extension = child.getAttr("Extension");
            let contentType = child.getAttr("ContentType");
            if (extension !== undefined && contentType !== undefined) {
                this.defaults.push({
                    extension: extension,
                    contentType: contentType
                });
            }
        });
        types?.visitChildren("Override", (child) => {
            let partName = child.getAttr("PartName");
            let contentType = child.getAttr("ContentType");
            if (partName !== undefined && contentType !== undefined) {
                this.overrides.push({
                    partName: partName,
                    contentType: contentType
                });
            }
        });
        return this;
    }
    toXml() {
        return Node.createDocument().appendChildren([
            Node.build("Types")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
                .appendChildren(this.defaults.map((def) => {
                return Node.build("Default")
                    .setAttr("Extension", def.extension)
                    .setAttr("ContentType", def.contentType);
            }))
                .appendChildren(this.overrides.map((override) => {
                return Node.build("Override")
                    .setAttr("PartName", override.partName)
                    .setAttr("ContentType", override.contentType);
            }))
        ]);
    }
    getContentTypeForExt(ext) {
        for (let def of this.defaults) {
            if (def.extension === ext)
                return def.contentType;
        }
        return null;
    }
    getOverrideForPartName(partName) {
        for (let override of this.overrides) {
            if (override.partName === partName)
                return override.contentType;
        }
        return null;
    }
    getContentTypeForPath(pathString) {
        pathString = normalizePath(pathString);
        let overrideContentType = this.getOverrideForPartName(pathString);
        if (overrideContentType !== null) {
            return overrideContentType;
        }
        const extension = path.extname(pathString).slice(1);
        return this.getContentTypeForExt(extension);
    }
    join(other) {
        for (let otherDef of other.defaults) {
            if (this.getContentTypeForExt(otherDef.extension) === null) {
                this.defaults.push({
                    ...otherDef
                });
            }
        }
        for (let otherOverride of other.overrides) {
            if (this.getOverrideForPartName(otherOverride.partName) === null) {
                this.overrides.push({
                    ...otherOverride
                });
            }
        }
    }
}

class Style extends Wrapper {
    getBaseStyle() {
        return getChildVal(this.node, "w:basedOn");
    }
    getLinkedStyle() {
        return getChildVal(this.node, "w:link");
    }
    getNextStyle() {
        return getChildVal(this.node, "w:link");
    }
    getName() {
        return getChildVal(this.node, "w:name");
    }
    getId() {
        return this.node.getAttr("w:styleId");
    }
    setBaseStyle(style) {
        setChildVal(this.node, "w:basedOn", style);
    }
    setLinkedStyle(style) {
        setChildVal(this.node, "w:link", style);
    }
    setNextStyle(style) {
        setChildVal(this.node, "w:next", style);
    }
    setName(name) {
        setChildVal(this.node, "w:name", name);
    }
    setId(id) {
        this.node.setAttr("w:styleId", id);
    }
}
class LatentStyles extends Wrapper {
    readOrCreate(node) {
        if (!node) {
            node = Node.build("w:latentStyles");
        }
        return this.readXml(node);
    }
    getLsdExceptions() {
        let result = new Map();
        this.node.visitChildren("w:lsdException", (child) => {
            let lsdException = new LSDException().readXml(child.shallowCopy());
            result.set(lsdException.name, lsdException);
        });
    }
}
class DocDefaults extends Wrapper {
    readOrCreate(node) {
        if (!node) {
            node = Node.build("w:docDefaults");
        }
        return this.readXml(node);
    }
}
class LSDException extends Wrapper {
    name;
    readXml(node) {
        this.name = node.getAttr("w:name");
        return this;
    }
    setName(name) {
        this.name = name;
        this.node.setAttr("w:name", name);
        return this;
    }
}
class Styles extends Serializable {
    styles = new Map();
    docDefaults = null;
    latentStyles = null;
    rels;
    readXml(xml) {
        this.styles = new Map();
        let styles = xml.getChild("w:styles");
        this.latentStyles = new LatentStyles().readOrCreate(styles.getChild("w:latentStyles"));
        this.docDefaults = new DocDefaults().readOrCreate(styles.getChild("w:docDefaults"));
        styles?.visitChildren("w:style", (child) => {
            let style = new Style().readXml(child.shallowCopy());
            this.styles.set(style.getId(), style);
        });
        return this;
    }
    toXml() {
        let styles = Array.from(this.styles.values());
        let result = Node.createDocument().appendChildren([
            Node.build("w:styles")
                .appendChildren([
                this.docDefaults.node.deepCopy(),
                this.latentStyles.node.deepCopy()
            ])
                .appendChildren(styles.map((style) => {
                return style.node.deepCopy();
            }))
        ]);
        result.getChild("w:styles").setAttrs(getProperXmlnsForDocument(result));
        return result;
    }
    removeStyle(style) {
        this.styles.delete(style.getId());
    }
    addStyle(style) {
        this.styles.set(style.getId(), style);
    }
    getStyleByName(name) {
        for (let [id, style] of this.styles) {
            if (style.getName() === name)
                return style;
        }
        return null;
    }
}

function getNodeLevel(node, tagName, index) {
    let result;
    let strIndex = String(index);
    node.visitChildren(tagName, (child) => {
        if (child.getAttr("w:ilvl") === strIndex) {
            result = child.shallowCopy();
            return false;
        }
    });
    if (!result) {
        result = Node.build(tagName).setAttr("w:ilvl", strIndex);
        node.appendChildren([result]);
    }
    return result;
}
class AbstractNum extends Wrapper {
    getId() {
        return this.node.getAttr("w:abstractNumId");
    }
    getLevel(index) {
        return getNodeLevel(this.node, "w:lvl", index);
    }
}
class Num extends Wrapper {
    getId() {
        return this.node.getAttr("w:numId");
    }
    getAbstractNumId() {
        return getChildVal(this.node, "w:abstractNumId");
    }
    getLevelOverride(index) {
        return getNodeLevel(this.node, "w:lvlOverride", index);
    }
    setId(id) {
        this.node.setAttr("w:numId", id);
    }
    setAbstractNumId(id) {
        setChildVal(this.node, "w:abstractNumId", id);
    }
}
class Numbering extends Serializable {
    abstractNums = new Map();
    nums = new Map();
    readXml(xml) {
        let styles = xml.getChild("w:numbering");
        styles?.visitChildren("w:abstractNum", (child) => {
            let abstractNum = new AbstractNum().readXml(child.shallowCopy());
            this.abstractNums.set(abstractNum.getId(), abstractNum);
        });
        styles?.visitChildren("w:num", (child) => {
            let num = new Num().readXml(child.shallowCopy());
            this.nums.set(num.getId(), num);
        });
        return this;
    }
    toXml() {
        let abstractNums = Array.from(this.abstractNums.values());
        let nums = Array.from(this.nums.values());
        return Node.createDocument().appendChildren([
            Node.build("w:numbering")
                .appendChildren(abstractNums.map((style) => {
                return style.node.deepCopy();
            }))
                .appendChildren(nums.map((style) => {
                return style.node.deepCopy();
            }))
        ]);
    }
    getUnusedNumId() {
        let index = 1;
        while (this.nums.has(String(index))) {
            index++;
        }
        return String(index);
    }
}

function getResourceTypeForMimeType(mimeType) {
    for (let key of Object.getOwnPropertyNames(resourceTypes)) {
        if (key === "__proto__")
            continue;
        if (resourceTypes[key].mimeType === mimeType) {
            return key;
        }
    }
}
const resourceTypes = {
    app: {
        mimeType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        xmlnsTag: "Properties"
    },
    core: {
        mimeType: "application/vnd.openxmlformats-package.core-properties+xml",
        xmlnsTag: "cp:coreProperties"
    },
    custom: {
        mimeType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        xmlnsTag: "Properties"
    },
    document: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        xmlnsTag: "w:document"
    },
    relationships: {
        mimeType: "application/vnd.openxmlformats-package.relationships+xml",
        xmlnsTag: "Relationships"
    },
    webSettings: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
        xmlnsTag: "webSettings"
    },
    numbering: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        xmlnsTag: "w:numbering"
    },
    settings: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        xmlnsTag: "w:settings"
    },
    styles: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        xmlnsTag: "w:styles"
    },
    fontTable: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
        xmlnsTag: "w:fonts"
    },
    theme: {
        mimeType: "application/vnd.openxmlformats-officedocument.theme+xml",
        xmlnsTag: "a:theme"
    },
    comments: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        xmlnsTag: "w:comments"
    },
    footnotes: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        xmlnsTag: "w:footnotes"
    },
    footer: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        xmlnsTag: "w:ftr"
    },
    header: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        xmlnsTag: "w:hdr"
    },
    png: {
        mimeType: "image/png"
    },
};

const contentTypesPath = "/[Content_Types].xml";
const globalRelsPath = "/_rels/.rels";
class WordResource {
    document;
    path;
    resource;
    rels = null;
    constructor(document, path, resource) {
        this.document = document;
        this.path = path;
        this.resource = resource;
    }
    saveRels() {
        if (!this.rels)
            return;
        let relsXml = this.rels.toXml();
        this.document.saveXml(getRelsPath(this.path), relsXml);
    }
    save() {
        let xml = this.resource.toXml();
        let contentType = this.document.contentTypes.resource.getContentTypeForPath(this.path);
        if (contentType) {
            let resourceType = getResourceTypeForMimeType(contentType);
            if (resourceType) {
                fixXmlns(xml, resourceTypes[resourceType].xmlnsTag);
            }
        }
        this.document.saveXml(this.path, xml);
        this.saveRels();
    }
    setRels(rels) {
        this.rels = rels;
        return this;
    }
}
const ResourceFactories = {
    generic: (document, path, xml) => {
        return new WordResource(document, path, new Wrapper().readXml(xml));
    },
    genericWithRel: (document, path, xml, rel) => {
        return new WordResource(document, path, new Wrapper().readXml(xml)).setRels(rel);
    },
    styles: (document, path, xml) => {
        return new WordResource(document, path, new Styles().readXml(xml));
    },
    numbering: (document, path, xml) => {
        return new WordResource(document, path, new Numbering().readXml(xml));
    },
    relationships: (document, path, xml) => {
        return new WordResource(document, path, new Relationships().readXml(xml));
    },
    contentTypes: (document, path, xml) => {
        return new WordResource(document, path, new ContentTypes().readXml(xml));
    },
};
class WordDocument {
    zipContents;
    wrappers = new Map();
    contentTypes;
    globalRels;
    numbering;
    styles;
    document;
    settings;
    fontTable;
    comments;
    headers = [];
    footers = [];
    async load(path) {
        const contents = await fs.promises.readFile(path);
        this.zipContents = await JSZip.loadAsync(contents);
        this.contentTypes = await this.createResourceForPath(ResourceFactories.contentTypes, contentTypesPath);
        this.globalRels = await this.createResourceForPath(ResourceFactories.relationships, globalRelsPath);
        this.document = await this.createResourceForType(ResourceFactories.genericWithRel, resourceTypes.document);
        this.styles = await this.createResourceForType(ResourceFactories.styles, resourceTypes.styles);
        this.settings = await this.createResourceForType(ResourceFactories.generic, resourceTypes.settings);
        this.numbering = await this.createResourceForType(ResourceFactories.numbering, resourceTypes.numbering);
        this.fontTable = await this.createResourceForType(ResourceFactories.generic, resourceTypes.fontTable);
        this.comments = await this.createResourceForType(ResourceFactories.generic, resourceTypes.comments);
        this.headers = await this.createResourcesForType(ResourceFactories.generic, resourceTypes.header);
        this.footers = await this.createResourcesForType(ResourceFactories.generic, resourceTypes.footer);
        return this;
    }
    getSinglePathForMimeType(type) {
        let paths = this.getPathsForMimeType(type);
        if (paths.length !== 1)
            return null;
        return paths[0];
    }
    async createResourceForType(factory, type) {
        let path = this.getSinglePathForMimeType(type.mimeType);
        if (!path)
            return null;
        return await this.createResourceForPath(factory, path);
    }
    async createResourcesForType(factory, type) {
        let paths = this.getPathsForMimeType(type.mimeType);
        return await Promise.all(paths.map(path => this.createResourceForPath(factory, path)));
    }
    async createResourceForPath(factory, pathString) {
        pathString = normalizePath(pathString);
        if (this.wrappers.has(pathString)) {
            throw new Error("This resource have already been created");
        }
        let relsPath = getRelsPath(pathString);
        let relationships = null;
        let relationshipsXml = await this.getXml(relsPath);
        if (relationshipsXml) {
            relationships = new Relationships().readXml(relationshipsXml);
        }
        let resource = factory(this, pathString, await this.getXml(pathString), relationships);
        this.wrappers.set(pathString, resource);
        return resource;
    }
    getPathsForMimeType(type) {
        let result = [];
        this.zipContents.forEach((path) => {
            let mimeType = this.contentTypes.resource.getContentTypeForPath(path);
            if (mimeType === type) {
                result.push(path);
            }
        });
        return result;
    }
    hasFile(path) {
        return this.zipContents.file(path) !== null;
    }
    async getFile(path) {
        return await this.zipContents.file(path.slice(1)).async("arraybuffer");
    }
    async getXml(path) {
        let contents = this.zipContents.file(path.slice(1));
        if (!contents)
            return null;
        return Node.fromXmlString(await contents.async("string"));
    }
    saveFile(path, data) {
        this.zipContents.file(path.slice(1), data);
    }
    saveXml(path, xml) {
        this.zipContents.file(path.slice(1), xml.toXmlString());
    }
    async save(path) {
        for (let [path, resource] of this.wrappers) {
            resource.save();
        }
        const contents = await this.zipContents.generateAsync({ type: "uint8array" });
        await fs.writeFileSync(path, contents);
    }
}

function* uniqueNameGenerator(name) {
    let index = 0;
    while (true) {
        let nameCandidate = name;
        if (index > 0)
            nameCandidate += "_" + index;
        yield nameCandidate;
        index++;
    }
}

class ParagraphTemplateSubstitution {
    document;
    template;
    replacement;
    setDocument(document) {
        this.document = document;
        return this;
    }
    setTemplate(template) {
        this.template = template;
        return this;
    }
    setReplacement(replacement) {
        this.replacement = replacement;
        return this;
    }
    perform() {
        const body = this.document.document.resource.toXml().getChild("w:document").getChild("w:body");
        this.replaceParagraphsWithTemplate(body);
        return this;
    }
    replaceParagraphsWithTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            let paragraphText = "";
            child.visitSubtree("w:t", (textNode) => {
                paragraphText += getTextContents(textNode);
            });
            if (paragraphText.indexOf(this.template) === -1) {
                continue;
            }
            if (paragraphText !== this.template) {
                throw new Error(`The ${this.template} pattern should be the only text of the paragraph`);
            }
            body.removeChild([i]);
            let replacement = this.replacement();
            body.insertChildren(replacement, [i]);
            i += replacement.length - 1;
        }
    }
}

const styleTags = ["w:pStyle", "w:rStyle", "w:tblStyle"];
class StyledTemplateSubstitution {
    // MARK: Options
    styleConversion = new Map();
    stylesToMigrate = new Set();
    listsConversion = {};
    source;
    target;
    template;
    allowUnrecognizedStyles = false;
    // MARK: Internal fields
    styleIdConversion = new Map();
    styleIdsToMigrate = new Set();
    transferredRels = new Map();
    transferredStyles = new Map();
    transferredResources = new Map();
    transferredNumbering = new Map();
    setStyleConversion(conversion) {
        this.styleConversion = conversion;
        return this;
    }
    setStylesToMigrate(styles) {
        this.stylesToMigrate = styles;
        return this;
    }
    setListConversion(conversion) {
        this.listsConversion = conversion;
        return this;
    }
    setSource(document) {
        this.source = document;
        return this;
    }
    setTarget(document) {
        this.target = document;
        return this;
    }
    setTemplate(template) {
        this.template = template;
        return this;
    }
    setAllowUnrecognizedStyles(allowUnrecognizedStyles) {
        this.allowUnrecognizedStyles = allowUnrecognizedStyles;
        return this;
    }
    async perform() {
        this.transferredRels.clear();
        this.transferredStyles.clear();
        this.updateStyleIds();
        let sourceBody = this.source.document.resource.toXml().getChild("w:document").getChild("w:body");
        let transferredBody = Node.build(sourceBody.getTagName());
        this.copySubtree(sourceBody, transferredBody);
        let promises = [];
        for (let [from, to] of this.transferredResources) {
            promises.push(this.transferResource(from, to));
        }
        await Promise.all(promises);
        new ParagraphTemplateSubstitution()
            .setDocument(this.target)
            .setTemplate(this.template)
            .setReplacement(() => transferredBody.getChildren())
            .perform();
        return this;
    }
    copySubtree(from, to) {
        from.visitChildren((child, index) => {
            if (child.isLeaf()) {
                let copy = child.deepCopy();
                to.insertChildren([copy], [index]);
                return false;
            }
            let tagName = child.getTagName();
            let attributes = child.getAttrs();
            let childCopy = Node.build(tagName).setAttrs(attributes);
            to.insertChildren([childCopy], [index]);
            this.copySubtree(child, childCopy);
            this.handleCopy(childCopy);
            return true;
        });
    }
    async transferResource(from, to) {
        this.target.saveFile(to, await this.source.getFile(from));
    }
    updateStyleIds() {
        this.styleIdConversion.clear();
        for (let [sourceStyleName, targetStyleName] of this.styleConversion) {
            let sourceStyle = this.source.styles.resource.getStyleByName(sourceStyleName);
            let targetStyle = this.target.styles.resource.getStyleByName(targetStyleName);
            if (!sourceStyle) {
                throw new Error("Could not find style named '" + sourceStyleName + "' in source document");
            }
            if (!targetStyle) {
                throw new Error("Could not find style named '" + targetStyleName + "' in target document");
            }
            this.styleIdConversion.set(sourceStyle.getId(), targetStyle.getId());
        }
        this.styleIdsToMigrate.clear();
        for (let styleName of this.stylesToMigrate) {
            let style = this.source.styles.resource.getStyleByName(styleName);
            this.styleIdsToMigrate.add(style.getId());
        }
        return this;
    }
    maybeCopyResource(pathString) {
        if (!this.transferredResources.has(pathString)) {
            // Choose unique path for the new resource
            let extname = path.extname(pathString);
            let basename = path.basename(pathString, extname);
            let dirname = path.dirname(pathString);
            for (let candidate of uniqueNameGenerator(dirname + "/" + basename)) {
                candidate += extname;
                if (!this.target.hasFile(candidate)) {
                    this.transferredResources.set(pathString, candidate);
                    break;
                }
            }
        }
        return this.transferredResources.get(pathString);
    }
    transferRelation(relId) {
        if (!this.source.document.rels) {
            throw new Error("Relation attribute is used in a resource without relationships xml");
        }
        if (!this.target.document.rels) {
            this.target.document.setRels(new Relationships());
        }
        let relation = this.source.document.rels.relations.get(relId);
        let existingRelation = this.target.document.rels.getRelForTarget(relation.target);
        if (existingRelation) {
            return existingRelation.id;
        }
        let dirname = path.dirname(this.source.document.path);
        let relationPath = normalizePath(dirname + "/" + relation.target);
        let resourcePath = this.maybeCopyResource(relationPath);
        let newRelation = {
            type: relation.type,
            target: path.relative(dirname, resourcePath),
            id: this.target.document.rels.getUnusedId()
        };
        this.target.document.rels.relations.set(newRelation.id, newRelation);
        this.transferredRels.set(relId, newRelation.id);
    }
    handleStyleNode(node) {
        if (styleTags.indexOf(node.getTagName()) === -1)
            return;
        let style = node.getAttr("w:val");
        let converted = this.getConvertedStyleId(style);
        if (converted !== null)
            node.setAttr("w:val", converted);
    }
    getTransferredNumberingFor(numPr) {
        let numId = numPr.getChild("w:numId").getAttr("w:val");
        let transferredNumbering = this.transferredNumbering.get(numId);
        if (transferredNumbering) {
            return transferredNumbering;
        }
        let docNumbering = this.source.numbering.resource;
        let numbering = docNumbering.nums.get(numId);
        let abstractNum = docNumbering.abstractNums.get(numbering.getAbstractNumId());
        let format = abstractNum.getLevel(0)
            .getChild("w:numFmt")
            ?.getAttr("w:val");
        let intendedNumbering;
        if (format === "decimal") {
            intendedNumbering = this.listsConversion.decimal;
        }
        else if (format === "bullet") {
            intendedNumbering = this.listsConversion.bullet;
        }
        else {
            throw new Error("Could not convert numbering with format " + format);
        }
        transferredNumbering = {
            styleName: this.target.styles.resource.getStyleByName(intendedNumbering.styleName).getId(),
            numId: this.createNumbering(intendedNumbering.numId)
        };
        this.transferredNumbering.set(numId, transferredNumbering);
        return transferredNumbering;
    }
    handleParagraphNumbering(node) {
        if (node.getTagName() !== "w:pPr")
            return;
        let numPr = node.getChild("w:numPr");
        if (!numPr)
            return;
        let transferredNumbering = this.getTransferredNumberingFor(numPr);
        numPr.getChild("w:numId").setAttr("w:val", transferredNumbering.numId);
        node.removeChildren("w:pStyle");
        node.pushChild(Node.build("w:pStyle")
            .setAttr("w:val", transferredNumbering.styleName));
    }
    createNumbering(abstractId) {
        let newNumId = this.target.numbering.resource.getUnusedNumId();
        let newNumbering = new Num().readXml(Node.build("w:num"));
        newNumbering.setAbstractNumId(abstractId);
        newNumbering.setId(newNumId);
        for (let i = 0; i < 8; i++) {
            newNumbering.getLevelOverride(i).appendChildren([
                Node.build("w:startOverride")
                    .setAttr("w:val", "1")
            ]);
        }
        this.target.numbering.resource.nums.set(newNumId, newNumbering);
        return newNumId;
    }
    getConvertedStyleId(styleId) {
        let conversion = this.styleIdConversion.get(styleId);
        if (conversion === undefined && this.styleIdsToMigrate.has(styleId)) {
            conversion = this.migrateSourceStyle(styleId);
        }
        if (conversion !== undefined) {
            return conversion;
        }
        if (this.allowUnrecognizedStyles) {
            return null;
        }
        // Some editors can break when the used style was not
        // defined. As an example, LibreOffice failed to render
        // the table properly because its paragraphs were using
        // a non-existent style.
        // To prevent this, the default behaviour is to throw
        // an error when unrecognized class is detected.
        this.reportUnrecognizedSourceStyle(styleId);
    }
    handleRelNode(node) {
        for (let attr of ["r:id", "r:embed"]) {
            let relId = node.getAttr(attr);
            if (!relId)
                continue;
            if (!this.transferredRels.has(relId)) {
                this.transferRelation(relId);
            }
            node.setAttr(attr, this.transferredRels.get(relId));
        }
    }
    handleCopy(node) {
        this.handleStyleNode(node);
        this.handleParagraphNumbering(node);
        this.handleRelNode(node);
    }
    getUnusedTargetStyleName(baseName) {
        for (let name of uniqueNameGenerator(baseName)) {
            if (!this.target.styles.resource.getStyleByName(name))
                return name;
        }
    }
    getUnusedTargetStyleId(baseName) {
        for (let name of uniqueNameGenerator("migrated_" + baseName)) {
            if (!this.source.styles.resource.styles.has(name))
                return name;
        }
    }
    migrateSourceStyle(styleId) {
        let style = this.source.styles.resource.styles.get(styleId);
        let copiedStyle = new Style().readXml(style.node.deepCopy());
        let copiedStyleId = this.getUnusedTargetStyleId(styleId);
        let copiedStyleName = this.getUnusedTargetStyleName(style.getName());
        this.styleIdConversion.set(styleId, copiedStyleId);
        copiedStyle.setId(copiedStyleId);
        copiedStyle.setName(copiedStyleName);
        let nextStyle = copiedStyle.getNextStyle();
        let baseStyle = copiedStyle.getBaseStyle();
        let linkedStyle = copiedStyle.getLinkedStyle();
        if (nextStyle !== null) {
            let converted = this.getConvertedStyleId(nextStyle);
            if (converted !== null)
                copiedStyle.setNextStyle(converted);
        }
        if (baseStyle !== null) {
            let converted = this.getConvertedStyleId(baseStyle);
            if (converted !== null)
                copiedStyle.setBaseStyle(converted);
        }
        if (linkedStyle !== null) {
            let converted = this.getConvertedStyleId(linkedStyle);
            if (converted !== null)
                copiedStyle.setLinkedStyle(converted);
        }
        this.target.styles.resource.addStyle(copiedStyle);
        return copiedStyleId;
    }
    reportUnrecognizedSourceStyle(styleId) {
        let style = this.source.styles.resource.styles.get(styleId);
        throw new Error("The source document contains style " + style.getName() + " which is not handled " +
            "by the style conversion ruleset. Provide the substitution for this style, mark it to be " +
            "migrated or use setAllowUnrecognizedStyle(true).");
    }
}

class InlineTemplateSubstitution {
    document;
    template;
    replacement;
    setDocument(document) {
        this.document = document;
        return this;
    }
    setTemplate(template) {
        this.template = template;
        return this;
    }
    setReplacement(replacement) {
        this.replacement = replacement;
        return this;
    }
    replaceInlineTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(keys.text, (textNode) => {
                    textNode.setText(textNode.getText().replace(this.template, this.replacement));
                });
            });
        }
    }
    removeParagraphsWithTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            let found = false;
            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(keys.text, (textNode) => {
                    let text = textNode.getText();
                    if (text.indexOf(this.template) !== null) {
                        found = true;
                    }
                });
                return !found;
            });
            if (found) {
                body.removeChild([i]);
                i--;
            }
        }
    }
    performIn(body) {
        if (this.replacement === "@none") {
            this.removeParagraphsWithTemplate(body);
        }
        else {
            this.replaceInlineTemplate(body);
        }
    }
    perform() {
        let document = this.document;
        let documentBody = document.document.resource.toXml().getChild("w:document").getChild("w:body");
        this.performIn(documentBody);
        for (let header of document.headers) {
            this.performIn(header.resource.toXml().getChild("w:hdr"));
        }
        for (let footer of document.footers) {
            this.performIn(footer.resource.toXml().getChild("w:ftr"));
        }
        return this;
    }
}

/*
    This code was taken from https://github.com/mvhenderson/pandoc-filter-node
    License for pandoc-filter-node:

    Copyright (c) 2014 Mike Henderson

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.
 */
function isElement(x) {
    return (typeof x === "object" && x && "t" in x) || false;
}
function walkPandocElement(object, action) {
    if (Array.isArray(object)) {
        let array = [];
        for (const element of object) {
            if (!isElement(element)) {
                array.push(walkPandocElement(element, action));
                continue;
            }
            let replacement = action(element);
            if (replacement) {
                if (Array.isArray(replacement)) {
                    array.push(...replacement);
                }
                else {
                    array.push(replacement);
                }
            }
            else {
                array.push(walkPandocElement(element, action));
            }
        }
        return array;
    }
    if (typeof object === "object" && object !== null) {
        let result = {};
        for (const key of Object.getOwnPropertyNames(object)) {
            if (key === "__proto__")
                continue;
            result[key] = walkPandocElement(object[key], action);
        }
        return result;
    }
    return object;
}
function metaElementToSource(value) {
    let result = [];
    walkPandocElement(value, (child) => {
        if (child.t === "Str")
            result.push(child.c);
        else if (child.t === "Strong")
            result.push("__" + metaElementToSource(child.c) + "__");
        else if (child.t === "Emph")
            result.push("_" + metaElementToSource(child.c) + "_");
        else if (child.t === "Space")
            result.push(" ");
        else if (child.t === "LineBreak")
            result.push("\n");
        else if (child.t === "Code")
            result.push("`" + child.c[1] + "`");
        else
            return undefined;
        return child;
    });
    return result.join("");
}

function getOpenxmlInjection(xml) {
    return {
        t: "RawBlock",
        c: ["openxml", xml]
    };
}
class PandocJsonPatcher {
    pandocJson;
    constructor(pandocJson) {
        this.pandocJson = pandocJson;
    }
    replaceDivWithClass(className, replacement) {
        this.pandocJson.blocks = walkPandocElement(this.pandocJson.blocks, (element) => {
            if (element.t === "Div" && element.c[0][1].indexOf(className) !== -1) {
                return getOpenxmlInjection(replacement(metaElementToSource(element)).toXmlString());
            }
        });
        return this;
    }
}

class PandocJsonMeta {
    section;
    path;
    constructor(section, path = "") {
        this.section = section;
        this.path = path;
    }
    getSection(path) {
        let any = this.getChild(path);
        return new PandocJsonMeta(any, this.getAbsPath(path));
    }
    asArray() {
        if (this.section === undefined) {
            this.reportNotExistError("", "MetaList");
        }
        else if (this.section.t !== "MetaList") {
            this.reportWrongTypeError("", "MetaList", this.section.t);
        }
        else {
            return this.section.c.map((element, index) => {
                return new PandocJsonMeta(element, this.getAbsPath(String(index)));
            });
        }
    }
    getString(path = "") {
        let child = this.getChild(path);
        if (!child) {
            this.reportNotExistError(path, "MetaInlines");
        }
        else if (child.t !== "MetaInlines") {
            this.reportWrongTypeError(path, "MetaInlines", child.t);
        }
        else {
            return metaElementToSource(child.c);
        }
    }
    reportNotExistError(relPath, expected) {
        let absPath = this.getAbsPath(relPath);
        throw new Error("Failed to parse document metadata: expected to have " + expected + " at path " + absPath);
    }
    reportWrongTypeError(relPath, expected, actual) {
        let absPath = this.getAbsPath(relPath);
        throw new Error("Failed to parse document metadata: expected " + expected + " at path " + absPath + ", got " +
            actual + " instead");
    }
    getAbsPath(relPath) {
        if (this.path.length) {
            if (relPath.length) {
                return this.path + "." + relPath;
            }
            return this.path;
        }
        return relPath;
    }
    getChild(path) {
        if (!path.length)
            return this.section;
        let result = this.section;
        for (let component of path.split(".")) {
            // Be safe from prototype pollution
            if (component === "__proto__")
                return undefined;
            if (!result)
                return undefined;
            if (result.t === "MetaMap") {
                result = result.c[component];
            }
            if (result.t === "MetaList") {
                let index = Number.parseInt(component);
                if (!Number.isNaN(index)) {
                    result = result.c[index];
                }
            }
        }
        return result;
    }
}

const pandocFlags = ["--tab-stop=8"];
const languages = ["ru", "en"];
const resourcesDir = path__namespace.dirname(process__namespace.argv[1]) + "/../resources";
function getLinksParagraphs(document, meta) {
    let styleId = document.styles.resource.getStyleByName("ispLitList").getId();
    let numId = "80";
    let links = meta.getSection("links").asArray();
    let result = [];
    for (let link of links) {
        let paragraph = buildParagraphWithStyle(styleId);
        let style = paragraph.getChild("w:pPr");
        style.pushChild(buildNumPr("0", numId));
        paragraph.pushChild(buildParagraphTextTag(link.getString()));
        result.push(paragraph);
    }
    return result;
}
function getAuthors(document, meta, language) {
    let styleId = document.styles.resource.getStyleByName("ispAuthor").getId();
    let authors = meta.getSection("authors").asArray();
    let result = [];
    let authorIndex = 1;
    for (let author of authors) {
        let paragraph = buildParagraphWithStyle(styleId);
        let name = author.getString("name_" + language);
        let orcid = author.getString("orcid");
        let email = author.getString("email");
        let indexLine = String(authorIndex);
        let authorLine = `${name}, ORCID: ${orcid}, <${email}>`;
        paragraph.pushChild(buildParagraphTextTag(indexLine, [buildSuperscriptTextStyle()]));
        paragraph.pushChild(buildParagraphTextTag(authorLine));
        result.push(paragraph);
        authorIndex++;
    }
    return result;
}
function getOrganizations(document, meta, language) {
    let styleId = document.styles.resource.getStyleByName("ispAuthor").getId();
    let organizations = meta.getSection("organizations_" + language).asArray();
    let orgIndex = 1;
    let result = [];
    for (let organization of organizations) {
        let paragraph = buildParagraphWithStyle(styleId);
        let indexLine = String(orgIndex);
        paragraph.pushChild(buildParagraphTextTag(indexLine, [buildSuperscriptTextStyle()]));
        paragraph.pushChild(buildParagraphTextTag(organization.getString()));
        result.push(paragraph);
        orgIndex++;
    }
    return result;
}
function getAuthorsDetail(document, meta) {
    let styleId = document.styles.resource.getStyleByName("Body Text").getId();
    let authors = meta.getSection("authors").asArray();
    let result = [];
    for (let author of authors) {
        for (let language of languages) {
            let line = author.getString("details_" + language);
            let newParagraph = buildParagraphWithStyle(styleId);
            newParagraph.pushChild(buildParagraphTextTag(line));
            result.push(newParagraph);
        }
    }
    return result;
}
function getImageCaption(document, content) {
    let styleId = document.styles.resource.getStyleByName("Image Caption").getId();
    return Node.build("w:p").appendChildren([
        Node.build("w:pPr").appendChildren([
            Node.build("w:pStyle").setAttr("w:val", styleId),
            Node.build("w:contextualSpacing").setAttr("w:val", "true"),
        ]),
        buildParagraphTextTag(content)
    ]);
}
function getListingCaption(document, content) {
    let styleId = document.styles.resource.getStyleByName("Body Text").getId();
    return Node.build("w:p").appendChildren([
        Node.build("w:pPr").appendChildren([
            Node.build("w:pStyle").setAttr("w:val", styleId),
            Node.build("w:jc").setAttr("w:val", "left"),
        ]),
        buildParagraphTextTag(content, [
            Node.build("w:i"),
            Node.build("w:iCs"),
            Node.build("w:sz").setAttr("w:val", "18"),
            Node.build("w:szCs").setAttr("w:val", "18"),
        ])
    ]);
}
function patchPandocJson(contentDoc, pandocJson) {
    new PandocJsonPatcher(pandocJson)
        .replaceDivWithClass("img-caption", (contents) => getImageCaption(contentDoc, contents))
        .replaceDivWithClass("table-caption", (contents) => getListingCaption(contentDoc, contents))
        .replaceDivWithClass("listing-caption", (contents) => getListingCaption(contentDoc, contents));
}
async function patchTemplateDocx(templateDoc, contentDoc, pandocJsonMeta) {
    await new StyledTemplateSubstitution()
        .setSource(contentDoc)
        .setTarget(templateDoc)
        .setTemplate("{{{body}}}")
        .setStyleConversion(new Map([
        ["Heading 1", "ispSubHeader-1 level"],
        ["Heading 2", "ispSubHeader-2 level"],
        ["Heading 3", "ispSubHeader-3 level"],
        ["Author", "ispAuthor"],
        ["Abstract Title", "ispAnotation"],
        ["Abstract", "ispAnotation"],
        ["Block Text", "ispText_main"],
        ["Body Text", "ispText_main"],
        ["First Paragraph", "ispText_main"],
        ["Normal", "Normal"],
        ["Compact", "Normal"],
        ["Source Code", "ispListing"],
        ["Verbatim Char", "ispListing Знак"],
        ["Image Caption", "ispPicture_sign"],
        ["Table", "Table Grid"]
    ]))
        .setStylesToMigrate(new Set([
        ...tokenClasses
    ]))
        .setAllowUnrecognizedStyles(false)
        .setListConversion({
        decimal: {
            styleName: "ispNumList",
            numId: "33"
        },
        bullet: {
            styleName: "ispList1",
            numId: "43"
        }
    })
        .perform();
    let inlineSubstitution = new InlineTemplateSubstitution().setDocument(templateDoc);
    let paragraphSubstitution = new ParagraphTemplateSubstitution().setDocument(templateDoc);
    for (let language of languages) {
        let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"];
        for (let template of templates) {
            let template_lang = template + "_" + language;
            let replacement = pandocJsonMeta.getString(template_lang);
            inlineSubstitution
                .setTemplate("{{{" + template_lang + "}}}")
                .setReplacement(replacement)
                .perform();
        }
        let header = pandocJsonMeta.getString("page_header_" + language);
        if (header === "@use_citation") {
            header = pandocJsonMeta.getString("for_citation_" + language);
        }
        inlineSubstitution
            .setTemplate("{{{page_header_" + language + "}}}")
            .setReplacement(header)
            .perform();
        paragraphSubstitution
            .setTemplate("{{{authors_" + language + "}}}")
            .setReplacement(() => getAuthors(templateDoc, pandocJsonMeta, language))
            .perform();
        paragraphSubstitution
            .setTemplate("{{{organizations_" + language + "}}}")
            .setReplacement(() => getOrganizations(templateDoc, pandocJsonMeta, language))
            .perform();
    }
    paragraphSubstitution
        .setTemplate("{{{links}}}")
        .setReplacement(() => getLinksParagraphs(templateDoc, pandocJsonMeta))
        .perform();
    paragraphSubstitution
        .setTemplate("{{{authors_detail}}}")
        .setReplacement(() => getAuthorsDetail(templateDoc, pandocJsonMeta))
        .perform();
}
async function main() {
    let argv = process__namespace.argv;
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>");
        process__namespace.exit(1);
    }
    let markdownSource = argv[2];
    let targetPath = argv[3];
    let tmpDocPath = targetPath + ".tmp";
    let contentDoc = await new WordDocument().load(tmpDocPath);
    let markdown = await fs__namespace.promises.readFile(markdownSource, "utf-8");
    let pandocJson = await markdownToPandocJson(markdown, pandocFlags);
    patchPandocJson(contentDoc, pandocJson);
    await pandocJsonToDocx(pandocJson, ["-o", tmpDocPath]);
    let pandocJsonMeta = new PandocJsonMeta(pandocJson.meta["ispras_templates"]);
    let templateDoc = await new WordDocument().load(resourcesDir + '/isp-reference.docx');
    await patchTemplateDocx(templateDoc, contentDoc, pandocJsonMeta);
    await templateDoc.save(targetPath);
}
main().then();

exports.languages = languages;
//# sourceMappingURL=main.js.map
