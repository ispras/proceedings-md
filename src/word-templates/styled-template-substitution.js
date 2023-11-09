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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.StyledTemplateSubstitution = void 0;
const XML = __importStar(require("src/xml"));
const OXML = __importStar(require("src/word/oxml"));
const path_1 = __importDefault(require("path"));
const relationships_1 = __importDefault(require("src/word/relationships"));
const styles_1 = require("src/word/styles");
const numbering_1 = require("src/word/numbering");
const utils_1 = require("src/utils");
const paragraph_template_substitution_1 = __importDefault(require("src/word-templates/paragraph-template-substitution"));
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
        let transferredBody = XML.Node.build(sourceBody.getTagName());
        this.copySubtree(sourceBody, transferredBody);
        let promises = [];
        for (let [from, to] of this.transferredResources) {
            promises.push(this.transferResource(from, to));
        }
        await Promise.all(promises);
        new paragraph_template_substitution_1.default()
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
            let childCopy = XML.Node.build(tagName).setAttrs(attributes);
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
            let extname = path_1.default.extname(pathString);
            let basename = path_1.default.basename(pathString, extname);
            let dirname = path_1.default.dirname(pathString);
            for (let candidate of (0, utils_1.uniqueNameGenerator)(dirname + "/" + basename)) {
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
            this.target.document.setRels(new relationships_1.default());
        }
        let relation = this.source.document.rels.relations.get(relId);
        let existingRelation = this.target.document.rels.getRelForTarget(relation.target);
        if (existingRelation) {
            return existingRelation.id;
        }
        let dirname = path_1.default.dirname(this.source.document.path);
        let relationPath = OXML.normalizePath(dirname + "/" + relation.target);
        let resourcePath = this.maybeCopyResource(relationPath);
        let newRelation = {
            type: relation.type,
            target: path_1.default.relative(dirname, resourcePath),
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
        node.pushChild(XML.Node.build("w:pStyle")
            .setAttr("w:val", transferredNumbering.styleName));
    }
    createNumbering(abstractId) {
        let newNumId = this.target.numbering.resource.getUnusedNumId();
        let newNumbering = new numbering_1.Num().readXml(XML.Node.build("w:num"));
        newNumbering.setAbstractNumId(abstractId);
        newNumbering.setId(newNumId);
        for (let i = 0; i < 8; i++) {
            newNumbering.getLevelOverride(i).appendChildren([
                XML.Node.build("w:startOverride")
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
        for (let name of (0, utils_1.uniqueNameGenerator)(baseName)) {
            if (!this.target.styles.resource.getStyleByName(name))
                return name;
        }
    }
    getUnusedTargetStyleId(baseName) {
        for (let name of (0, utils_1.uniqueNameGenerator)("migrated_" + baseName)) {
            if (!this.source.styles.resource.styles.has(name))
                return name;
        }
    }
    migrateSourceStyle(styleId) {
        let style = this.source.styles.resource.styles.get(styleId);
        let copiedStyle = new styles_1.Style().readXml(style.node.deepCopy());
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
exports.StyledTemplateSubstitution = StyledTemplateSubstitution;
//# sourceMappingURL=styled-template-substitution.js.map