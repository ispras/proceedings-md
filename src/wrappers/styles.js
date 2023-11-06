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
exports.LSDException = exports.DocDefaults = exports.LatentStyles = exports.Style = void 0;
const XML = __importStar(require("../xml.js"));
class Style extends XML.Wrapper {
    getBaseStyle() {
        return this.getChildVal("w:basedOn");
    }
    getLinkedStyle() {
        return this.getChildVal("w:link");
    }
    getNextStyle() {
        return this.getChildVal("w:link");
    }
    getName() {
        return this.getChildVal("w:name");
    }
    getId() {
        return this.node.getAttr("w:styleId");
    }
    setBaseStyle(style) {
        this.setChildVal("w:basedOn", style);
    }
    setLinkedStyle(style) {
        this.setChildVal("w:link", style);
    }
    setNextStyle(style) {
        this.setChildVal("w:next", style);
    }
    setName(name) {
        this.setChildVal("w:name", name);
    }
    setId(id) {
        this.node.setAttr("w:styleId", id);
    }
    getChildVal(tag) {
        let basedOnTag = this.node.getChild(tag);
        if (basedOnTag)
            return basedOnTag.getAttr("w:val");
        return null;
    }
    setChildVal(tag, value) {
        if (value === null) {
            this.node.removeChildren(tag);
        }
        else {
            let basedOnTag = this.node.getChild(tag);
            if (basedOnTag)
                basedOnTag.setAttr("w:val", value);
            else
                this.node.appendChildren([
                    XML.Node.build(tag).setAttr("w:val", value)
                ]);
        }
    }
}
exports.Style = Style;
class LatentStyles extends XML.Wrapper {
    readOrCreate(node) {
        if (!node) {
            node = XML.Node.build("w:latentStyles");
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
exports.LatentStyles = LatentStyles;
class DocDefaults extends XML.Wrapper {
    readOrCreate(node) {
        if (!node) {
            node = XML.Node.build("w:docDefaults");
        }
        return this.readXml(node);
    }
}
exports.DocDefaults = DocDefaults;
class LSDException extends XML.Wrapper {
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
exports.LSDException = LSDException;
class Styles extends XML.Serializable {
    styles = new Map();
    docDefaults = null;
    latentStyles = null;
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
        return XML.Node.createDocument().appendChildren([
            XML.Node.build("w:styles")
                .setAttr("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
                .setAttr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                .setAttr("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                .setAttr("xmlns:w14", "http://schemas.microsoft.com/office/word/2010/wordml")
                .setAttr("xmlns:w15", "http://schemas.microsoft.com/office/word/2012/wordml")
                .setAttr("mc:Ignorable", "w14 w15")
                .appendChildren([
                this.docDefaults.node.deepCopy(),
                this.latentStyles.node.deepCopy()
            ])
                .appendChildren(styles.map((style) => {
                return style.node.deepCopy();
            }))
        ]);
    }
    removeStyle(style) {
        this.styles.delete(style.getId());
    }
    addStyle(style) {
        this.styles.set(style.getId(), style);
    }
}
exports.default = Styles;
//# sourceMappingURL=styles.js.map