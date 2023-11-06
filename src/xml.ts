import {XMLBuilder, XMLParser} from "fast-xml-parser";

export const keys = {
    comment: "__comment__",
    text: "__text__",
    attributes: ":@",
    document: "__document__"
}

export const parser = new XMLParser({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true,
    trimValues: false,
    commentPropName: keys.comment,
    textNodeName: keys.text
})

export const builder = new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    preserveOrder: true,
    commentPropName: keys.comment,
    textNodeName: keys.text
})

function checkFilter(filter: Filter, node: Node) {
    if (!filter) return true
    if (typeof filter === "string") {
        return node.getTagName() === filter
    }
    return filter(node)
}

export type Filter = string | ((node: Node) => boolean) | null
export type VisitCallback<PathType> = (node: Node, path: PathType) => (boolean | void)
export type Path = number[]

function getVisitArgs<PathType>(args: any[]) {
    let filter: Filter = null
    let callback: VisitCallback<PathType> = args[0]
    let startPosition: PathType | undefined = args[1]

    if (typeof args[1] === "function") {
        filter = args[0]
        callback = args[1]
        startPosition = args[2]
    }

    return {
        filter: filter,
        callback: callback,
        startPosition: startPosition
    }
}

export class Node {
    private element: any
    private tempDestroyed: boolean = false

    constructor(element: any) {
        if (Array.isArray(element)) {
            throw new Error("XML.Node must be constructed from the xml object, not its children list")
        }
        this.element = element
    }

    getTagName() {
        this.checkTemporary()
        for (let key of Object.getOwnPropertyNames(this.element)) {
            // Be safe from prototype pollution
            if (key === "__proto__" || key === keys.attributes) continue
            return key
        }
        return null
    }

    pushChild(child: Node) {
        this.checkTemporary()

        let tagName = this.getTagName()
        if (!Array.isArray(this.element[tagName])) {
            throw new Error("Cannot call pushChild on " + tagName + " element")
        }

        let childArray = this.element[tagName] as any[]
        childArray.push(child.raw())
    }

    unshiftChild(child: Node) {
        this.checkTemporary()

        let tagName = this.getTagName()
        if (!Array.isArray(this.element[tagName])) {
            throw new Error("Cannot call pushChild on " + tagName + " element")
        }

        let childArray = this.element[tagName] as any[]
        childArray.unshift(child.raw())
    }

    getChildren(filter: Filter | null = null): Node[] {
        this.checkTemporary()

        let result = []
        this.visitChildren(filter, (child) => {
            result.push(child.shallowCopy())
        })
        return result
    }

    getChild(filter: Filter | null): Node | null
    getChild(path: Path): Node | null
    getChild(arg: Filter | Path | null = null): Node | null {
        this.checkTemporary()

        if (Array.isArray(arg)) {
            let path = arg as Path
            if (path.length === 0) {
                return this
            }

            let result = new Node(this.element)

            for (let i = 0; i < path.length; i++) {
                if (!result.element) return null
                let tagName = result.getTagName()
                let pathComponent = path[i]
                let children = result.element[tagName]

                if (pathComponent < 0) {
                    result.element = children[children.length + pathComponent]
                } else {
                    result.element = children[pathComponent]
                }
            }

            if (!result.element) return null
            return result
        } else {
            let filter = arg as Filter
            let result = null
            this.visitChildren(filter, (child) => {
                if (result) {
                    throw new Error("Element have multiple children matching the given filter")
                }
                result = child.shallowCopy()
            })
            return result
        }
    }

    visitChildren(filter: Filter, callback: VisitCallback<number>, startIndex?: number): void
    visitChildren(callback: VisitCallback<number>, startIndex?: number): void
    visitChildren(...args: any[]) {
        this.checkTemporary()

        let {
            filter,
            callback,
            startPosition
        } = getVisitArgs<number>(args)

        let tagName = this.getTagName()
        if (!Array.isArray(this.element[tagName])) {
            return
        }

        let index = startPosition ?? 0

        let tmpNode = new Node(null)
        for (let child of this.element[tagName]) {
            tmpNode.element = child
            if (checkFilter(filter, tmpNode)) {
                if (callback(tmpNode, index) === false) {
                    break
                }
            }
            index++
        }

        tmpNode.markDestroyed()
    }

    visitSubtree(filter: Filter, callback: VisitCallback<Path>, startPath?: Path): void
    visitSubtree(callback: VisitCallback<Path>, startPath?: Path): void
    visitSubtree(...args: any[]) {
        this.checkTemporary()

        let {
            filter,
            callback,
            startPosition
        } = getVisitArgs<Path>(args)

        let tmpNode = new Node(null)
        let startPath = startPosition ?? []
        let startDepth = 0

        let path = []

        const walk = (node: Node) => {
            let tagName = node.getTagName()
            let children = node.element[tagName]

            if (!Array.isArray(children)) {
                return
            }

            let depth = path.length
            let startIndex = 0
            if (depth < startDepth && startPath.length) {
                startIndex = startPath[startPath.length]
                startDepth = depth
            }

            for (let index = startIndex; index < children.length; index++) {
                path.push(index)
                tmpNode.element = children[index]
                let filterPass = checkFilter(filter, tmpNode)
                let goDeeper = true

                if (filterPass) {
                    goDeeper = callback(tmpNode, path) === true
                }

                if (goDeeper) {
                    walk(tmpNode)
                }

                path.pop()
            }
        }

        walk(this)

        tmpNode.markDestroyed()
    }

    removeChild(path: Path) {
        let topIndex = path.pop()

        let child = this.getChild(path)
        let tagName = child.getTagName()
        child.element[tagName].splice(topIndex, 1)

        path.push(topIndex)
    }

    removeChildren(filter: Filter = null) {
        this.checkTemporary()
        let tagName = this.getTagName()

        if (!Array.isArray(this.element[tagName])) {
            return;
        }

        let node = new Node(null)

        for (let i = 0; i < this.element[tagName].length; i++) {
            node.element = this.element[tagName][i]
            if (checkFilter(filter, node)) {
                this.element[tagName].splice(i, 1)
                i--
            }
        }

        node.markDestroyed()
    }

    isTextNode() {
        this.checkTemporary()

        return this.getTagName() == keys.text
    }

    isCommentNode() {
        this.checkTemporary()

        return this.getTagName() == keys.comment
    }

    getText(): string {
        this.checkTemporary()

        if (!this.isTextNode()) {
            throw new Error("getText() is called on " + this.getTagName() + " element")
        }
        return String(this.element[keys.text])
    }

    setText(text: string) {
        this.checkTemporary()

        if (!this.isTextNode()) {
            throw new Error("setText() is called on " + this.getTagName() + " element")
        }
        this.element[keys.text] = text
    }

    getComment(): string {
        this.checkTemporary()

        if (!this.isCommentNode()) {
            throw new Error("getComment() is called on " + this.getTagName() + " element")
        }

        let textChild = this.getChild(keys.text)
        return textChild.getText()
    }

    static build(tagName: string) {
        let element = {}
        element[tagName] = []
        return new Node(element)
    }

    static createDocument(args: { [key: string]: string } = {}) {
        args = Object.assign({
            version: "1.0",
            encoding: "UTF-8",
            standalone: "yes"
        }, args)
        let document = this.build(keys.document)
        document.appendChildren([
            Node.build("?xml")
                .setAttrs(args)
                .appendChildren([
                    Node.buildTextNode("")
                ])
        ])
        return document
    }

    static buildTextNode(text: string) {
        let element = {}
        element[keys.text] = text
        return new Node(element)
    }

    setAttr(attribute: string, value: string) {
        this.checkTemporary()

        if (!this.element[keys.attributes]) {
            this.element[keys.attributes] = {}
        }

        this.element[keys.attributes][attribute] = value
        return this
    }

    setAttrs(attributes: { [key: string]: string }) {
        this.checkTemporary()

        this.element[keys.attributes] = attributes
        return this
    }

    getAttr(attribute: string): string {
        this.checkTemporary()

        if (!this.element[keys.attributes]) {
            return undefined
        }
        return String(this.element[keys.attributes][attribute])
    }

    clearChildren(path: Path = []) {
        this.checkTemporary()

        let parent = this.getChild(path)
        parent.element[parent.getTagName()] = []
        return this
    }

    insertChildren(children: Node[], path: Path) {
        this.checkTemporary()

        let insertIndex = path.pop()
        let parent = this.getChild(path)
        path.push(insertIndex)

        let lastChildren = parent.element[parent.getTagName()]
        if (insertIndex < 0) {
            insertIndex = children.length + insertIndex + 1
        }

        lastChildren.splice(insertIndex, 0, ...children.map(child => child.raw()))
        return this
    }

    appendChildren(children: Node[], path: Path = []) {
        path.push(-1)
        this.insertChildren(children, path)
        path.pop()
        return this
    }

    unshiftChildren(children: Node[], path: Path = []) {
        path.push(0)
        this.insertChildren(children, path)
        path.pop()
        return this
    }

    assign(another: Node) {
        this.checkTemporary()

        if (this === another) {
            return
        }

        if (this.element) {
            this.element[this.getTagName()] = undefined
        } else {
            this.element = {}
        }

        this.element[another.getTagName()] = JSON.parse(JSON.stringify(another.raw()[another.getTagName()]))

        if (another.raw()[keys.attributes]) {
            this.element[keys.attributes] = JSON.parse(JSON.stringify(another.raw()[keys.attributes]))
        } else {
            this.element[keys.attributes] = {}
        }

        return this
    }

    static fromXmlString(str: string) {
        let object = parser.parse(str)
        let wrapped = {}
        wrapped[keys.document] = object
        return new Node(wrapped)
    }

    toXmlString() {
        this.checkTemporary()
        if (this.getTagName() === keys.document) {
            return builder.build(this.element[keys.document])
        } else {
            return builder.build([this.element])
        }
    }

    raw() {
        this.checkTemporary()
        return this.element
    }

    private checkTemporary() {
        if (this.tempDestroyed) {
            throw new Error("Method access to an outdated temporary Node. Make sure to call .shallowCopy() on temporary " +
                "nodes before accessing them outside your visitChildren/visitSubtree body scope")
        }
    }

    private markDestroyed() {
        // From now on, the checkTemporary method will throw
        this.tempDestroyed = true
    }

    shallowCopy() {
        this.checkTemporary()

        return new Node(this.element)
    }

    deepCopy() {
        return new Node(null).assign(this)
    }
}

export class Serializable {
    readXmlString(xmlString: string): this {
        this.readXml(Node.fromXmlString(xmlString))
        return this
    }

    readXml(xml: Node): this {
        throw new Error("readXml is not implemented")
        return this
    }

    toXmlString() {
        return this.toXml().toXmlString()
    }

    toXml(): Node {
        throw new Error("toXml is not implemented")
    }
}

export class Wrapper extends Serializable {
    node: Node | null = null

    readXml(xml: Node): this {
        this.node = xml
        return this
    }

    toXml() {
        return this.node
    }
}
