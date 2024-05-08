import * as XML from "src/xml";
import * as OXML from "src/word/oxml";

function getNodeLevel(node: XML.Node, tagName: string, index: number) {
    let result: XML.Node | null

    let strIndex = String(index)

    node.visitChildren(tagName, (child) => {
        if(child.getAttr("w:ilvl") === strIndex) {
            result = child.shallowCopy()
            return false
        }
    })

    if(!result) {
        result = XML.Node.build(tagName).setAttr("w:ilvl", strIndex)
        node.appendChildren([result])
    }

    return result
}

export class AbstractNum extends XML.Wrapper {
    getId() {
        return this.node.getAttr("w:abstractNumId")
    }

    getLevel(index: number) {
        return getNodeLevel(this.node, "w:lvl", index)
    }
}

export class Num extends XML.Wrapper {
    getId() {
        return this.node.getAttr("w:numId")
    }

    getAbstractNumId() {
        return OXML.getChildVal(this.node, "w:abstractNumId")
    }

    getLevelOverride(index: number) {
        return getNodeLevel(this.node, "w:lvlOverride", index)
    }

    setId(id: string) {
        this.node.setAttr("w:numId", id)
    }

    setAbstractNumId(id: string) {
        OXML.setChildVal(this.node, "w:abstractNumId", id)
    }
}

export default class Numbering extends XML.Serializable {
    abstractNums: Map<string, AbstractNum> = new Map()
    nums: Map<string, Num> = new Map()

    readXml(xml: XML.Node) {

        let styles = xml.getChild("w:numbering")

        styles?.visitChildren("w:abstractNum", (child) => {
            let abstractNum = new AbstractNum().readXml(child.shallowCopy())
            this.abstractNums.set(abstractNum.getId(), abstractNum)
        })

        styles?.visitChildren("w:num", (child) => {
            let num = new Num().readXml(child.shallowCopy())
            this.nums.set(num.getId(), num)
        })

        return this
    }

    toXml() {
        let abstractNums = Array.from(this.abstractNums.values())
        let nums = Array.from(this.nums.values())

        return XML.Node.createDocument().appendChildren([
            XML.Node.build("w:numbering")
                .appendChildren(abstractNums.map((style) => {
                    return style.node.deepCopy()
                }))
                .appendChildren(nums.map((style) => {
                    return style.node.deepCopy()
                }))
        ])
    }

    getUnusedNumId() {
        let index = 1
        while(this.nums.has(String(index))) {
            index++
        }
        return String(index)
    }
}