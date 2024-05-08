import * as XML from 'src/xml.js'
import * as OXML from 'src/word/oxml'
import Relationships from "./relationships";

export class Style extends XML.Wrapper {
    getBaseStyle() {
        return OXML.getChildVal(this.node, "w:basedOn")
    }

    getLinkedStyle() {
        return OXML.getChildVal(this.node, "w:link")
    }

    getNextStyle() {
        return OXML.getChildVal(this.node, "w:link")
    }

    getName() {
        return OXML.getChildVal(this.node, "w:name")
    }

    getId() {
        return this.node.getAttr("w:styleId")
    }

    setBaseStyle(style: string | null) {
        OXML.setChildVal(this.node, "w:basedOn", style)
    }

    setLinkedStyle(style: string | null) {
        OXML.setChildVal(this.node, "w:link", style)
    }

    setNextStyle(style: string | null) {
        OXML.setChildVal(this.node, "w:next", style)
    }

    setName(name: string) {
        OXML.setChildVal(this.node, "w:name", name)
    }

    setId(id: string) {
        this.node.setAttr("w:styleId", id)
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
    rels: Relationships

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

        let result = XML.Node.createDocument().appendChildren([
            XML.Node.build("w:styles")
                .appendChildren([
                    this.docDefaults.node.deepCopy(),
                    this.latentStyles.node.deepCopy()
                ])
                .appendChildren(styles.map((style) => {
                    return style.node.deepCopy()
                }))
        ])

        result.getChild("w:styles").setAttrs(OXML.getProperXmlnsForDocument(result))

        return result
    }

    removeStyle(style: Style) {
        this.styles.delete(style.getId())
    }

    addStyle(style: Style) {
        this.styles.set(style.getId(), style)
    }

    getStyleByName(name: string) {
        for(let [id, style] of this.styles) {
            if(style.getName() === name) return style
        }
        return null
    }
}