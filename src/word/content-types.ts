import * as XML from 'src/xml.js'
import * as OXML from 'src/word/oxml'
import path from "path";

export interface ContentTypeDefault {
    extension: string,
    contentType: string
}

export interface ContentTypeOverride {
    partName: string
    contentType: string
}

export default class ContentTypes extends XML.Serializable {
    defaults: ContentTypeDefault[]
    overrides: ContentTypeOverride[]

    readXml(xml: XML.Node) {
        this.defaults = []
        this.overrides = []

        let types = xml.getChild("Types")

        types?.visitChildren("Default", (child) => {
            let extension = child.getAttr("Extension")
            let contentType = child.getAttr("ContentType")

            if (extension !== undefined && contentType !== undefined) {
                this.defaults.push({
                    extension: extension,
                    contentType: contentType
                })
            }
        })

        types?.visitChildren("Override", (child) => {
            let partName = child.getAttr("PartName")
            let contentType = child.getAttr("ContentType")

            if (partName !== undefined && contentType !== undefined) {
                this.overrides.push({
                    partName: partName,
                    contentType: contentType
                })
            }
        })

        return this
    }

    toXml() {
        return XML.Node.createDocument().appendChildren([
            XML.Node.build("Types")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
                .appendChildren(this.defaults.map((def) => {
                    return XML.Node.build("Default")
                        .setAttr("Extension", def.extension)
                        .setAttr("ContentType", def.contentType)
                }))
                .appendChildren(this.overrides.map((override) => {
                    return XML.Node.build("Override")
                        .setAttr("PartName", override.partName)
                        .setAttr("ContentType", override.contentType)
                }))
        ])
    }

    getContentTypeForExt(ext: string) {
        for(let def of this.defaults) {
            if(def.extension === ext) return def.contentType
        }
        return null
    }

    getOverrideForPartName(partName: string) {
        for(let override of this.overrides) {
            if(override.partName === partName) return override.contentType
        }
        return null
    }

    getContentTypeForPath(pathString: string) {
        pathString = OXML.normalizePath(pathString)
        let overrideContentType = this.getOverrideForPartName(pathString)

        if(overrideContentType !== null) {
            return overrideContentType
        }

        const extension = path.extname(pathString).slice(1)
        return this.getContentTypeForExt(extension)
    }

    join(other: ContentTypes) {
        for(let otherDef of other.defaults) {
            if(this.getContentTypeForExt(otherDef.extension) === null) {
                this.defaults.push({
                    ...otherDef
                })
            }
        }

        for(let otherOverride of other.overrides) {
            if(this.getOverrideForPartName(otherOverride.partName) === null) {
                this.overrides.push({
                    ...otherOverride
                })
            }
        }
    }
}