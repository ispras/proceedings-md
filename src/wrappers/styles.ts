import * as XML from '../xml.js'

export class Style extends XML.Wrapper {
    getBaseStyle() {
        return this.getChildVal("w:basedOn")
    }

    getLinkedStyle() {
        return this.getChildVal("w:link")
    }

    getNextStyle() {
        return this.getChildVal("w:link")
    }

    getName() {
        return this.getChildVal("w:name")
    }

    getId() {
        return this.node.getAttr("w:styleId")
    }

    setBaseStyle(style: string | null) {
        this.setChildVal("w:basedOn", style)
    }

    setLinkedStyle(style: string | null) {
        this.setChildVal("w:link", style)
    }

    setNextStyle(style: string | null) {
        this.setChildVal("w:next", style)
    }

    setName(name: string) {
        this.setChildVal("w:name", name)
    }

    setId(id: string) {
        this.node.setAttr("w:styleId", id)
    }

    private getChildVal(tag: string) {
        let basedOnTag = this.node.getChild(tag)
        if(basedOnTag) return basedOnTag.getAttr("w:val")
        return null
    }

    private setChildVal(tag: string, value: string | null) {
        if(value === null) {
            this.node.removeChildren(tag)
        } else {
            let basedOnTag = this.node.getChild(tag)
            if(basedOnTag) basedOnTag.setAttr("w:val", value)
            else this.node.appendChildren([
                XML.Node.build(tag).setAttr("w:val", value)
            ])
        }
    }
}

export class LatentStyles extends XML.Wrapper {
    readOrCreate(node: XML.Node): this {
        if(!node) {
            node = XML.Node.build("w:latentStyles")
        }
        return this.readXml(node)
    }

    getLsdExceptions() {
        let result = new Map()
        this.node.visitChildren("w:lsdException", (child) => {
            let lsdException = new LSDException().readXml(child.shallowCopy())
            result.set(lsdException.name, lsdException)
        })
    }
}

export class DocDefaults extends XML.Wrapper {
    readOrCreate(node: XML.Node): this {
        if(!node) {
            node = XML.Node.build("w:docDefaults")
        }
        return this.readXml(node)
    }
}

export class LSDException extends XML.Wrapper {
    name: string

    readXml(node: XML.Node) {
        this.name = node.getAttr("w:name")
        return this
    }

    setName(name: string) {
        this.name = name
        this.node.setAttr("w:name", name)
        return this
    }
}

export default class Styles extends XML.Serializable {
    styles: Map<string, Style> = new Map()
    docDefaults: DocDefaults | null = null
    latentStyles: LatentStyles | null = null

    readXml(xml: XML.Node) {
        this.styles = new Map()

        let styles = xml.getChild("w:styles")

        this.latentStyles = new LatentStyles().readOrCreate(styles.getChild("w:latentStyles"))
        this.docDefaults = new DocDefaults().readOrCreate(styles.getChild("w:docDefaults"))

        styles?.visitChildren("w:style", (child) => {
            let style = new Style().readXml(child.shallowCopy())

            this.styles.set(style.getId(), style)
        })

        return this
    }

    toXml() {
        let styles = Array.from(this.styles.values())

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
                    return style.node.deepCopy()
                }))
        ])
    }

    removeStyle(style: Style) {
        this.styles.delete(style.getId())
    }

    addStyle(style: Style) {
        this.styles.set(style.getId(), style)
    }
}