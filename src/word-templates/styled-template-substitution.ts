import * as XML from "src/xml"
import * as OXML from "src/word/oxml";
import path from "path"
import WordDocument from "src/word/word-document";
import Relationships, {Relation} from "src/word/relationships";
import {Style} from "src/word/styles";
import {Num} from "src/word/numbering";
import {uniqueNameGenerator} from "src/utils";
import ParagraphTemplateSubstitution from "src/word-templates/paragraph-template-substitution";

export interface ListConversion {
    styleName: string,
    numId: string
}

export interface ListsConversion {
    bullet?: ListConversion
    decimal?: ListConversion
}

const styleTags = ["w:pStyle", "w:rStyle", "w:tblStyle"]

export class StyledTemplateSubstitution {
    // MARK: Options
    private styleConversion: Map<string, string> = new Map()
    private stylesToMigrate: Set<string> = new Set()
    private listsConversion: ListsConversion = {}
    private source?: WordDocument
    private target?: WordDocument
    private template?: string
    private allowUnrecognizedStyles: boolean = false

    // MARK: Internal fields
    private styleIdConversion: Map<string, string> = new Map<string, string>()
    private styleIdsToMigrate: Set<string> = new Set()

    private transferredRels: Map<string, string> = new Map()
    private transferredStyles: Map<string, string> = new Map()
    private transferredResources: Map<string, string> = new Map()
    private transferredNumbering: Map<string, ListConversion> = new Map()

    setStyleConversion(conversion: Map<string, string>) {
        this.styleConversion = conversion
        return this
    }

    setStylesToMigrate(styles: Set<string>) {
        this.stylesToMigrate = styles
        return this
    }

    setListConversion(conversion: ListsConversion) {
        this.listsConversion = conversion
        return this
    }

    setSource(document: WordDocument) {
        this.source = document
        return this
    }

    setTarget(document: WordDocument) {
        this.target = document
        return this
    }

    setTemplate(template: string) {
        this.template = template
        return this
    }

    setAllowUnrecognizedStyles(allowUnrecognizedStyles: boolean) {
        this.allowUnrecognizedStyles = allowUnrecognizedStyles
        return this
    }

    async perform() {
        this.transferredRels.clear()
        this.transferredStyles.clear()
        this.updateStyleIds()

        let sourceBody = this.source.document.resource.toXml().getChild("w:document").getChild("w:body")

        let transferredBody = XML.Node.build(sourceBody.getTagName())
        this.copySubtree(sourceBody, transferredBody)

        let promises = []

        for (let [from, to] of this.transferredResources) {
            promises.push(this.transferResource(from, to))
        }

        await Promise.all(promises)

        new ParagraphTemplateSubstitution()
            .setDocument(this.target)
            .setTemplate(this.template)
            .setReplacement(() => transferredBody.getChildren())
            .perform()

        return this
    }

    private copySubtree(from: XML.Node, to: XML.Node) {
        from.visitChildren((child, index) => {
            if (child.isLeaf()) {
                let copy = child.deepCopy()
                to.insertChildren([copy], [index])
                return false
            }

            let tagName = child.getTagName()
            let attributes = child.getAttrs()

            let childCopy = XML.Node.build(tagName).setAttrs(attributes)
            to.insertChildren([childCopy], [index])

            this.copySubtree(child, childCopy)
            this.handleCopy(childCopy)

            return true
        })
    }

    private async transferResource(from: string, to: string) {
        this.target.saveFile(to, await this.source.getFile(from))
    }

    private updateStyleIds() {
        this.styleIdConversion.clear()

        for (let [sourceStyleName, targetStyleName] of this.styleConversion) {
            let sourceStyle = this.source.styles.resource.getStyleByName(sourceStyleName)
            let targetStyle = this.target.styles.resource.getStyleByName(targetStyleName)

            if (!sourceStyle) {
                throw new Error("Could not find style named '" + sourceStyleName + "' in source document")
            }

            if (!targetStyle) {
                throw new Error("Could not find style named '" + targetStyleName + "' in target document")
            }

            this.styleIdConversion.set(sourceStyle.getId(), targetStyle.getId())
        }

        this.styleIdsToMigrate.clear()

        for (let styleName of this.stylesToMigrate) {
            let style = this.source.styles.resource.getStyleByName(styleName)
            this.styleIdsToMigrate.add(style.getId())
        }

        return this
    }

    private maybeCopyResource(pathString: string) {
        if (!this.transferredResources.has(pathString)) {
            // Choose unique path for the new resource
            let extname = path.extname(pathString)
            let basename = path.basename(pathString, extname)
            let dirname = path.dirname(pathString)

            for (let candidate of uniqueNameGenerator(dirname + "/" + basename)) {
                candidate += extname
                if (!this.target.hasFile(candidate)) {
                    this.transferredResources.set(pathString, candidate)
                    break
                }
            }
        }

        return this.transferredResources.get(pathString)
    }

    private transferRelation(relId: string) {
        if (!this.source.document.rels) {
            throw new Error("Relation attribute is used in a resource without relationships xml")
        }

        if (!this.target.document.rels) {
            this.target.document.setRels(new Relationships())
        }

        let relation = this.source.document.rels.relations.get(relId)

        let existingRelation = this.target.document.rels.getRelForTarget(relation.target)

        if (existingRelation) {
            return existingRelation.id
        }

        let dirname = path.dirname(this.source.document.path)
        let relationPath = OXML.normalizePath(dirname + "/" + relation.target)
        let resourcePath = this.maybeCopyResource(relationPath)

        let newRelation: Relation = {
            type: relation.type,
            target: path.relative(dirname, resourcePath),
            id: this.target.document.rels.getUnusedId()
        }

        this.target.document.rels.relations.set(newRelation.id, newRelation)
        this.transferredRels.set(relId, newRelation.id)
    }

    private handleStyleNode(node: XML.Node) {
        if (styleTags.indexOf(node.getTagName()) === -1) return

        let style = node.getAttr("w:val")
        let converted = this.getConvertedStyleId(style)
        if (converted !== null) node.setAttr("w:val", converted)
    }

    private getTransferredNumberingFor(numPr: XML.Node) {
        let numId = numPr.getChild("w:numId").getAttr("w:val")
        let transferredNumbering = this.transferredNumbering.get(numId)

        if(transferredNumbering) {
            return transferredNumbering
        }

        let docNumbering = this.source.numbering.resource
        let numbering = docNumbering.nums.get(numId)
        let abstractNum = docNumbering.abstractNums.get(numbering.getAbstractNumId())

        let format = abstractNum.getLevel(0)
            .getChild("w:numFmt")
            ?.getAttr("w:val")
        let intendedNumbering: ListConversion

        if (format === "decimal") {
            intendedNumbering = this.listsConversion.decimal
        } else if (format === "bullet") {
            intendedNumbering = this.listsConversion.bullet
        } else {
            throw new Error("Could not convert numbering with format " + format)
        }

        transferredNumbering = {
            styleName: this.target.styles.resource.getStyleByName(intendedNumbering.styleName).getId(),
            numId: this.createNumbering(intendedNumbering.numId)
        }

        this.transferredNumbering.set(numId, transferredNumbering)
        return transferredNumbering
    }

    private handleParagraphNumbering(node: XML.Node) {
        if (node.getTagName() !== "w:pPr") return

        let numPr = node.getChild("w:numPr")
        if (!numPr) return

        let transferredNumbering = this.getTransferredNumberingFor(numPr)

        numPr.getChild("w:numId").setAttr("w:val", transferredNumbering.numId)

        node.removeChildren("w:pStyle")
        node.pushChild(
            XML.Node.build("w:pStyle")
                .setAttr("w:val", transferredNumbering.styleName)
        )

    }

    private createNumbering(abstractId: string) {
        let newNumId = this.target.numbering.resource.getUnusedNumId()
        let newNumbering = new Num().readXml(XML.Node.build("w:num"))
        newNumbering.setAbstractNumId(abstractId)
        newNumbering.setId(newNumId)

        for (let i = 0; i < 8; i++) {
            newNumbering.getLevelOverride(i).appendChildren([
                XML.Node.build("w:startOverride")
                    .setAttr("w:val", "1")
            ])
        }

        this.target.numbering.resource.nums.set(newNumId, newNumbering)
        return newNumId
    }


    private getConvertedStyleId(styleId: string) {
        let conversion = this.styleIdConversion.get(styleId)

        if (conversion === undefined && this.styleIdsToMigrate.has(styleId)) {
            conversion = this.migrateSourceStyle(styleId)
        }

        if (conversion !== undefined) {
            return conversion
        }

        if (this.allowUnrecognizedStyles) {
            return null
        }

        // Some editors can break when the used style was not
        // defined. As an example, LibreOffice failed to render
        // the table properly because its paragraphs were using
        // a non-existent style.
        // To prevent this, the default behaviour is to throw
        // an error when unrecognized class is detected.

        this.reportUnrecognizedSourceStyle(styleId)
    }

    private handleRelNode(node: XML.Node) {
        for (let attr of ["r:id", "r:embed"]) {
            let relId = node.getAttr(attr);
            if (!relId) continue

            if (!this.transferredRels.has(relId)) {
                this.transferRelation(relId)
            }

            node.setAttr(attr, this.transferredRels.get(relId));
        }
    }

    private handleCopy(node: XML.Node) {
        this.handleStyleNode(node)
        this.handleParagraphNumbering(node)
        this.handleRelNode(node)
    }

    private getUnusedTargetStyleName(baseName: string) {
        for (let name of uniqueNameGenerator(baseName)) {
            if (!this.target.styles.resource.getStyleByName(name)) return name
        }
    }

    private getUnusedTargetStyleId(baseName: string) {
        for (let name of uniqueNameGenerator("migrated_" + baseName)) {
            if (!this.source.styles.resource.styles.has(name)) return name
        }
    }

    private migrateSourceStyle(styleId: string) {
        let style = this.source.styles.resource.styles.get(styleId)

        let copiedStyle = new Style().readXml(style.node.deepCopy())
        let copiedStyleId = this.getUnusedTargetStyleId(styleId)
        let copiedStyleName = this.getUnusedTargetStyleName(style.getName())

        this.styleIdConversion.set(styleId, copiedStyleId)

        copiedStyle.setId(copiedStyleId)
        copiedStyle.setName(copiedStyleName)

        let nextStyle = copiedStyle.getNextStyle()
        let baseStyle = copiedStyle.getBaseStyle()
        let linkedStyle = copiedStyle.getLinkedStyle()

        if (nextStyle !== null) {
            let converted = this.getConvertedStyleId(nextStyle)
            if (converted !== null) copiedStyle.setNextStyle(converted)
        }

        if (baseStyle !== null) {
            let converted = this.getConvertedStyleId(baseStyle)
            if (converted !== null) copiedStyle.setBaseStyle(converted)
        }

        if (linkedStyle !== null) {
            let converted = this.getConvertedStyleId(linkedStyle)
            if (converted !== null) copiedStyle.setLinkedStyle(converted)
        }

        this.target.styles.resource.addStyle(copiedStyle)

        return copiedStyleId
    }

    private reportUnrecognizedSourceStyle(styleId: string) {
        let style = this.source.styles.resource.styles.get(styleId)
        throw new Error("The source document contains style " + style.getName() + " which is not handled " +
            "by the style conversion ruleset. Provide the substitution for this style, mark it to be " +
            "migrated or use setAllowUnrecognizedStyle(true).")
    }
}