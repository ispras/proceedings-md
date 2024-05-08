import JSZip from "jszip";
import * as XML from "src/xml"
import * as OXML from "src/word/oxml"
import fs from "fs";
import Relationships from "src/word/relationships";
import ContentTypes from "src/word/content-types";
import Styles from "src/word/styles";
import Numbering from "src/word/numbering";
import {getResourceTypeForMimeType, resourceTypes, WordResourceType} from "./resource-types";

export const contentTypesPath = "/[Content_Types].xml"
export const globalRelsPath = "/_rels/.rels"

export class WordResource<T extends XML.Serializable> {
    document: WordDocument
    path: string
    resource: T
    rels?: Relationships = null

    constructor(document: WordDocument, path: string, resource: T) {
        this.document = document
        this.path = path
        this.resource = resource
    }

    saveRels() {
        if(!this.rels) return
        let relsXml = this.rels.toXml()
        this.document.saveXml(OXML.getRelsPath(this.path), relsXml)
    }

    save() {
        let xml = this.resource.toXml()
        let contentType = this.document.contentTypes.resource.getContentTypeForPath(this.path)

        if(contentType) {
            let resourceType = getResourceTypeForMimeType(contentType)
            if (resourceType) {
                OXML.fixXmlns(xml, resourceTypes[resourceType].xmlnsTag)
            }
        }

        this.document.saveXml(this.path, xml)
        this.saveRels()
    }

    setRels(rels: Relationships): this {
        this.rels = rels
        return this
    }
}

export type WordXMLFactory<T extends XML.Serializable> = (document: WordDocument, path: string, xml: XML.Node, rels: Relationships) => WordResource<T>

export const ResourceFactories = {
    generic: (document, path, xml) => {
        return new WordResource(document, path, new XML.Wrapper().readXml(xml))
    },
    genericWithRel: (document, path, xml, rel) => {
        return new WordResource(document, path, new XML.Wrapper().readXml(xml)).setRels(rel)
    },
    styles: (document, path, xml) => {
        return new WordResource(document, path, new Styles().readXml(xml))
    },
    numbering: (document, path, xml) => {
        return new WordResource(document, path, new Numbering().readXml(xml))
    },
    relationships: (document, path, xml) => {
        return new WordResource(document, path, new Relationships().readXml(xml))
    },
    contentTypes: (document, path, xml) => {
        return new WordResource(document, path, new ContentTypes().readXml(xml))
    },
}

export default class WordDocument {

    zipContents: JSZip
    wrappers: Map<string, WordResource<XML.Serializable>> = new Map()

    contentTypes: WordResource<ContentTypes>
    globalRels: WordResource<Relationships>
    numbering: WordResource<Numbering>
    styles: WordResource<Styles>

    document: WordResource<XML.Serializable>
    settings: WordResource<XML.Serializable>
    fontTable: WordResource<XML.Serializable>
    comments: WordResource<XML.Serializable>

    headers: WordResource<XML.Serializable>[] = []
    footers: WordResource<XML.Serializable>[] = []

    async load(path: string) {
        const contents = await fs.promises.readFile(path)
        this.zipContents = await JSZip.loadAsync(contents)

        this.contentTypes = await this.createResourceForPath(ResourceFactories.contentTypes, contentTypesPath)
        this.globalRels = await this.createResourceForPath(ResourceFactories.relationships, globalRelsPath)
        this.document = await this.createResourceForType(ResourceFactories.genericWithRel, resourceTypes.document)
        this.styles = await this.createResourceForType(ResourceFactories.styles, resourceTypes.styles)
        this.settings = await this.createResourceForType(ResourceFactories.generic, resourceTypes.settings)
        this.numbering = await this.createResourceForType(ResourceFactories.numbering, resourceTypes.numbering)
        this.fontTable = await this.createResourceForType(ResourceFactories.generic, resourceTypes.fontTable)
        this.comments = await this.createResourceForType(ResourceFactories.generic, resourceTypes.comments)

        this.headers = await this.createResourcesForType(ResourceFactories.generic, resourceTypes.header)
        this.footers = await this.createResourcesForType(ResourceFactories.generic, resourceTypes.footer)

        return this
    }

    getSinglePathForMimeType(type: string) {
        let paths = this.getPathsForMimeType(type)
        if(paths.length !== 1) return null
        return paths[0]
    }

    async createResourceForType<T extends XML.Serializable>(factory: WordXMLFactory<T>, type: WordResourceType): Promise<WordResource<T>> {
        let path = this.getSinglePathForMimeType(type.mimeType)
        if(!path) return null
        return await this.createResourceForPath(factory, path)
    }

    async createResourcesForType<T extends XML.Serializable>(factory: WordXMLFactory<T>, type: WordResourceType): Promise<WordResource<T>[]> {
        let paths = this.getPathsForMimeType(type.mimeType)

        return await Promise.all(paths.map(path => this.createResourceForPath(factory, path)))
    }

    async createResourceForPath<T extends XML.Serializable>(factory: WordXMLFactory<T>, pathString: string): Promise<WordResource<T>> {
        pathString = OXML.normalizePath(pathString)

        if(this.wrappers.has(pathString)) {
            throw new Error("This resource have already been created")
        }

        let relsPath = OXML.getRelsPath(pathString)
        let relationships: Relationships | null = null

        let relationshipsXml = await this.getXml(relsPath)
        if(relationshipsXml) {
            relationships = new Relationships().readXml(relationshipsXml)
        }

        let resource = factory(this, pathString, await this.getXml(pathString), relationships)

        this.wrappers.set(pathString, resource)

        return resource
    }

    getPathsForMimeType(type: string) {
        let result = []
        this.zipContents.forEach((path) => {
            let mimeType = this.contentTypes.resource.getContentTypeForPath(path)
            if(mimeType === type) {
                result.push(path)
            }
        })
        return result
    }

    hasFile(path: string) {
        return this.zipContents.file(path) !== null
    }

    async getFile(path: string) {
        return await this.zipContents.file(path.slice(1)).async("arraybuffer")
    }

    async getXml(path: string) {
        let contents = this.zipContents.file(path.slice(1))
        if(!contents) return null
        return XML.Node.fromXmlString(await contents.async("string"))
    }

    saveFile(path: string, data: ArrayBuffer) {
        this.zipContents.file(path.slice(1), data)
    }

    saveXml(path: string, xml: XML.Node) {
        this.zipContents.file(path.slice(1), xml.toXmlString())
    }

    async save(path: string) {
        for(let [path, resource] of this.wrappers) {
            resource.save()
        }

        const contents = await this.zipContents.generateAsync({type: "uint8array"})
        await fs.writeFileSync(path, contents);
    }
}
