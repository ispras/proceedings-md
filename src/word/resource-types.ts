
export interface WordResourceType {
    mimeType: string
    xmlnsTag?: string
}

export function getResourceTypeForMimeType(mimeType: string) {
    for(let key of Object.getOwnPropertyNames(resourceTypes)) {
        if(key === "__proto__") continue
        if(resourceTypes[key].mimeType === mimeType) {
            return key
        }
    }
}

export const resourceTypes: {
    [key: string]: WordResourceType
} = {
    app: {
        mimeType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        xmlnsTag: "Properties"
    },
    core: {
        mimeType: "application/vnd.openxmlformats-package.core-properties+xml",
        xmlnsTag: "cp:coreProperties"
    },
    custom: {
        mimeType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        xmlnsTag: "Properties"
    },
    document: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        xmlnsTag: "w:document"
    },
    relationships: {
        mimeType: "application/vnd.openxmlformats-package.relationships+xml",
        xmlnsTag: "Relationships"
    },
    webSettings: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
        xmlnsTag: "webSettings"
    },
    numbering: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        xmlnsTag: "w:numbering"
    },
    settings: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        xmlnsTag: "w:settings"
    },
    styles: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        xmlnsTag: "w:styles"
    },
    fontTable: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
        xmlnsTag: "w:fonts"
    },
    theme: {
        mimeType: "application/vnd.openxmlformats-officedocument.theme+xml",
        xmlnsTag: "a:theme"
    },
    comments: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        xmlnsTag: "w:comments"
    },
    footnotes: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        xmlnsTag: "w:footnotes"
    },
    footer: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        xmlnsTag: "w:ftr"
    },
    header: {
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        xmlnsTag: "w:hdr"
    },
    png: {
        mimeType: "image/png"
    },
}