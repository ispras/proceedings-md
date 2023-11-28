import * as XML from 'src/xml.js'

export interface Relation {
    id: string
    type: string
    target: string
}

export default class Relationships extends XML.Serializable {
    relations: Map<string, Relation> = new Map()

    readXml(xml: XML.Node) {
        this.relations = new Map()

        xml.getChild("Relationships")?.visitChildren("Relationship", (child) => {
            let id = child.getAttr("Id")
            let type = child.getAttr("Type")
            let target = child.getAttr("Target")

            if (id !== undefined && type !== undefined && target !== undefined) {
                this.relations.set(id, {
                    id: id,
                    type: type,
                    target: target
                })
            }
        })

        return this
    }

    toXml() {
        let relations = Array.from(this.relations.values())

        return XML.Node.createDocument().appendChildren([
            XML.Node.build("Relationships")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
                .appendChildren(relations.map((relation) => {
                    return XML.Node.build("Relationship")
                        .setAttr("Id", relation.id)
                        .setAttr("Type", relation.type)
                        .setAttr("Target", relation.target)
                }))
        ])
    }

    getUnusedId() {
        let prefix = "rId"
        let index = 1

        while (this.relations.has(prefix + index)) {
            index++
        }

        return prefix + index
    }

    getRelForTarget(target: string) {
        for(let rel of this.relations.values()) {
            if(rel.target === target) {
                return rel
            }
        }
    }
}