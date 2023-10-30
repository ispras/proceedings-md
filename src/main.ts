import * as path from 'path';
import * as fs from 'fs';
import * as JSZip from 'jszip';
import {XMLParser, XMLBuilder, XMLValidator} from 'fast-xml-parser';
import {spawn} from "child_process";
import pandoc from "./pandoc";
import * as XML from "./xml";
import * as OXML from "./oxml";

const pandocFlags = ["--tab-stop=8"]

const properDocXmlns = new Map<string, string>([
    ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
    ["xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
    ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["xmlns:o", "urn:schemas-microsoft-com:office:office"],
    ["xmlns:v", "urn:schemas-microsoft-com:vml"],
    ["xmlns:w10", "urn:schemas-microsoft-com:office:word"],
    ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
    ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
    ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
])

let tagsWithRelId = new Map<string, string>([
    ["w:headerReference", "r:id"],
    ["w:footerReference", "r:id"],
    ["w:hyperlink", "r:id"],
    ["v:imagedata", "r:id"],
    ["a:blip", "r:embed"],
])

export const languages = ["ru", "en"]

function getStyleCrossReferences(styles: any): any[] {
    let result = []
    for (let style of XML.getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"]) continue
        result.push(style[XML.keys.attributes])

        let basedOnTag = XML.getChildTag(style["w:style"], "w:basedOn")
        if (basedOnTag) result.push(basedOnTag[XML.keys.attributes])

        let linkTag = XML.getChildTag(style["w:style"], "w:link")
        if (linkTag) result.push(linkTag[XML.keys.attributes])

        let nextTag = XML.getChildTag(style["w:style"], "w:next")
        if (nextTag) result.push(nextTag[XML.keys.attributes])
    }
    return result
}

function getDocStyleUseReferences(doc: any, result: any[] = [], met = new Set()): any[] {
    if (!doc || typeof doc !== "object" || met.has(doc)) {
        return result
    }
    met.add(doc)

    if (Array.isArray(doc)) {
        for (let child of doc) {
            result = getDocStyleUseReferences(child, result, met)
        }
    }

    let tagName = XML.getTagName(doc)
    if (tagName === "w:pStyle" || tagName == "w:rStyle") {
        result.push(doc[XML.keys.attributes])
    }
    result = getDocStyleUseReferences(doc[tagName], result, met)

    return result
}

function extractStyleDefs(styles: any): any[] {
    let result = []
    for (let style of XML.getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"]) continue

        if (style[XML.keys.attributes]["w:styleId"].startsWith("template-")) {
            let copy = JSON.parse(JSON.stringify(style))
            result.push(copy)
        }
    }
    return result
}

function patchStyleDefinitions(doc: any, styles: any, map: Map<string, string>) {
    let crossReferences = getStyleCrossReferences(styles)

    for (let ref of crossReferences) {
        if (ref["w:styleId"] && map.has(ref["w:styleId"])) {
            ref["w:styleId"] = map.get(ref["w:styleId"])
        }
    }
}

function patchStyleUseReferences(doc: any, styles: any, map: Map<string, string>) {
    let docReferences = getDocStyleUseReferences(doc)
    let crossReferences = getStyleCrossReferences(styles)

    for (let ref of docReferences.concat(crossReferences)) {
        if (ref["w:val"] && map.has(ref["w:val"])) {
            ref["w:val"] = map.get(ref["w:val"])
        }
    }
}

function getUsedStyles(doc: any): Set<string> {
    let references = getDocStyleUseReferences(doc)
    let set = new Set<string>()

    for (let ref of references) {
        set.add(ref["w:val"])
    }

    return set
}

function populateStyles(styles: Set<string>, table: Map<string, any>) {
    for (let styleId of styles) {
        let style = table.get(styleId)

        if (!style) {
            throw new Error("Style id " + styleId + " not found")
        }

        let basedOnTag = XML.getChildTag(style["w:style"], "w:basedOn")
        if (basedOnTag) styles.add(basedOnTag[XML.keys.attributes]["w:val"])

        let linkTag = XML.getChildTag(style["w:style"], "w:link")
        if (linkTag) styles.add(linkTag[XML.keys.attributes]["w:val"])

        let nextTag = XML.getChildTag(style["w:style"], "w:next")
        if (nextTag) styles.add(nextTag[XML.keys.attributes]["w:val"])
    }
}

function getUsedStylesDeep(doc: any, styleTable: Map<string, any>, requiredStyles: string[] = []): Set<string> {
    let usedStyles = getUsedStyles(doc)

    for (let requiredStyle of requiredStyles) {
        usedStyles.add(requiredStyle)
    }

    do {
        let size = usedStyles.size
        populateStyles(usedStyles, styleTable)
        if (usedStyles.size == size) break;
    } while (true);

    return usedStyles
}

function getStyleTable(styles: any): Map<string, any> {
    let table = new Map<string, any>()

    for (let style of XML.getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"]) continue
        table.set(style[XML.keys.attributes]["w:styleId"], style)
    }

    return table
}

function getStyleIdsByNameFromDefs(styles: any): Map<string, any> {
    let table = new Map<string, any>()

    for (let style of styles) {
        if (!style["w:style"]) continue
        let nameNode = XML.getChildTag(style["w:style"], "w:name")

        if (nameNode) {
            table.set(nameNode[XML.keys.attributes]["w:val"], style[XML.keys.attributes]["w:styleId"])
        }
    }

    return table
}

function addCollisionPatch(mappingTable: Map<string, string>, styleId: string): string {
    let newId = "template-" + mappingTable.size.toString()
    mappingTable.set(styleId, newId)
    return newId
}

function getMappingTable(usedStyles: Set<string>): Map<string, string> {
    let mappingTable = new Map<string, string>
    for (let style of usedStyles) {
        addCollisionPatch(mappingTable, style);
    }

    return mappingTable
}

function appendStyles(target, defs) {
    let styles = XML.getChildTag(target, "w:styles")["w:styles"]
    for (let def of defs) {
        styles.push(def)
    }
}

interface ListStyles {
    BulletList: NumIdPatchEntry
    OrderedList: NumIdPatchEntry

    [key: string]: NumIdPatchEntry | undefined
}

interface NumIdPatchEntry {
    styleName: string
    numId: string
}

function applyListStyles(doc, styles: ListStyles): Map<string, string> {

    let stack = []
    let currentState = undefined

    let met = new Set()
    let newStyles = new Map<string, string>()
    let lastId = 10000

    const walk = (doc) => {

        if (!doc || typeof doc !== "object" || met.has(doc)) {
            return
        }
        met.add(doc)

        for (let key of Object.getOwnPropertyNames(doc)) {
            walk(doc[key])

            if (key === "w:pPr" && currentState) {
                // Remove any old pStyle and add our own

                for (let i = 0; i < doc[key].length; i++) {
                    if (doc[key][i]["w:pStyle"]) {
                        doc[key].splice(i, 1)
                        i--
                    }
                }

                doc[key].unshift({
                    "w:pStyle": {},
                    ...XML.buildAttributes({"w:val": styles[currentState.listStyle].styleName})
                })
            }

            if (key === "w:numId" && currentState) {
                doc[XML.keys.attributes]["w:val"] = String(currentState.numId)
            }

            if (key === XML.keys.comment) {
                let commentValue = doc[key][0][XML.keys.text]
                // Switch between ordered list and bullet list
                // if comment is detected

                if (commentValue.indexOf("ListMode OrderedList") != -1) {
                    stack.push(currentState)
                    currentState = {
                        numId: lastId++,
                        listStyle: "OrderedList"
                    }
                    newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId)
                }

                if (commentValue.indexOf("ListMode BulletList") != -1) {
                    stack.push(currentState)
                    currentState = {
                        numId: lastId++,
                        listStyle: "BulletList"
                    }
                    newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId)
                }

                if (commentValue.indexOf("ListMode None") != -1) {
                    currentState = stack[stack.length - 1]
                    stack.pop()
                }
            }
        }
    }

    walk(doc)

    return newStyles
}

function removeCollidedStyles(styles: any, collisions: Set<string>) {
    let ignored = 0
    let newContents = []

    for (let style of XML.getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"] || !collisions.has(style[XML.keys.attributes]["w:styleId"])) {
            newContents.push(style)
        }
    }

    XML.getChildTag(styles, "w:styles")["w:styles"] = newContents
}

function copyLatentStyles(source, target) {
    let sourceStyles = XML.getChildTag(source, "w:styles")["w:styles"]
    let targetStyles = XML.getChildTag(target, "w:styles")["w:styles"]

    let sourceLatentStyles = XML.getChildTag(sourceStyles, "w:latentStyles")
    let targetLatentStyles = XML.getChildTag(targetStyles, "w:latentStyles")

    targetLatentStyles["w:latentStyles"] = JSON.parse(JSON.stringify(sourceLatentStyles["w:latentStyles"]))
    if (targetLatentStyles[XML.keys.attributes]) {
        targetLatentStyles[XML.keys.attributes] = JSON.parse(JSON.stringify(sourceLatentStyles[XML.keys.attributes]))
    }
}

function copyDocDefaults(source, target) {
    let sourceStyles = XML.getChildTag(source, "w:styles")["w:styles"]
    let targetStyles = XML.getChildTag(target, "w:styles")["w:styles"]

    let sourceDocDefaults = XML.getChildTag(sourceStyles, "w:docDefaults")
    let targetDocDefaults = XML.getChildTag(targetStyles, "w:docDefaults")

    targetDocDefaults["w:docDefaults"] = JSON.parse(JSON.stringify(sourceDocDefaults["w:docDefaults"]))
    if (sourceDocDefaults[XML.keys.attributes]) {
        targetDocDefaults[XML.keys.attributes] = JSON.parse(JSON.stringify(sourceDocDefaults[XML.keys.attributes]))
    }
}

async function copyFile(source, target, path) {
    target.file(path, await source.file(path).async("arraybuffer"))
}

function addNewNumberings(targetNumberingParsed: any, newListStyles: Map<string, string>) {
    let numberingTag = XML.getChildTag(targetNumberingParsed, "w:numbering")["w:numbering"]

    // <w:num w:numId="newNum">
    //   <w:abstractNumId w:val="oldNum"/>
    // </w:num>

    for (let [newNum, oldNum] of newListStyles) {

        let overrides = []
        for (let i = 0; i < 9; i++) {
            overrides.push({
                "w:lvlOverride": [{
                    "w:startOverride": [],
                    ...XML.buildAttributes({"w:val": "1"})
                }],
                ...XML.buildAttributes({"w:ilvl": String(i)})
            })
        }

        numberingTag.push({
            "w:num": [{
                "w:abstractNumId": [],
                ...XML.buildAttributes({"w:val": oldNum})
            }, ...overrides],
            ...XML.buildAttributes({"w:numId": newNum})
        })
    }
}

function addContentType(contentTypes, partName, contentType) {
    let typesTag = XML.getChildTag(contentTypes, "Types")["Types"]

    typesTag.push({
        "Override": [],
        ...XML.buildAttributes({
            "PartName": partName,
            "ContentType": contentType
        })
    })
}

function transferRels(source, target): Map<string, string> {
    let sourceRels = XML.getChildTag(source, "Relationships")["Relationships"]
    let targetRels = XML.getChildTag(target, "Relationships")["Relationships"]

    let presentIds = new Map<string, string>()
    let idMap = new Map<string, string>()

    for (let rel of targetRels) {
        presentIds.set(rel[XML.keys.attributes]["Target"], rel[XML.keys.attributes]["Id"])
    }

    let newIdCounter = 0

    for (let rel of sourceRels) {
        if (presentIds.has(rel[XML.keys.attributes]["Target"])) {
            idMap.set(rel[XML.keys.attributes]["Id"], presentIds.get(rel[XML.keys.attributes]["Target"]))
        } else {
            let newId = "template-id-" + (newIdCounter++)
            let relCopy = JSON.parse(JSON.stringify(rel))
            relCopy[XML.keys.attributes]["Id"] = newId
            targetRels.push(relCopy)
            idMap.set(rel[XML.keys.attributes]["Id"], newId)
        }
    }

    return idMap
}

function getRawText(tag): string {
    let result = ""
    let tagName = XML.getTagName(tag)

    if (tagName === XML.keys.text) {
        result += tag[XML.keys.text]
    }
    if (Array.isArray(tag[tagName])) {
        for (let child of tag[tagName]) {
            result += getRawText(child)
        }
    }

    return result
}

function replaceInlineTemplate(body: any[], template: string, value: string) {
    if (value === "@none") {
        let i = findParagraphWithPattern(body, template, 0);
        for (; i !== null; i = findParagraphWithPattern(body, template, i)) {
            body.splice(i, 1)
            i = i - 1;
        }
    } else {
        replaceStringTemplate(body, template, value)
    }
}

function replaceStringTemplate(tag: any, template: string, value: string) {
    if (Array.isArray(tag)) {
        for (let child of tag) {
            replaceStringTemplate(child, template, value)
        }
        return
    }

    let tagName = XML.getTagName(tag)

    if (tagName === XML.keys.text) {
        tag[XML.keys.text] = String(tag[XML.keys.text]).replace(template, value)
    } else if (typeof tag[tagName] === "object") {
        replaceStringTemplate(tag[tagName], template, value)
    }
}

function getParagraphText(paragraph: any): string {
    let result = ""

    if (paragraph["w:t"]) {
        result += getRawText(paragraph)
    }

    for (let name of Object.getOwnPropertyNames(paragraph)) {
        if (name === XML.keys.attributes) {
            continue
        }
        if (Array.isArray(paragraph[name])) {
            for (let child of paragraph[name]) {
                result += getParagraphText(child)
            }
        }
    }

    return result
}

function findParagraphWithPattern(body: any, pattern: string, startIndex: number = 0): number | null {
    for (let i = startIndex; i < body.length; i++) {
        let text = getParagraphText(body[i])
        if (text.indexOf(pattern) == -1) {
            continue
        }
        return i
    }

    return null
}

function findParagraphWithPatternStrict(body: any, pattern: string, startIndex: number = 0): number | null {
    let paragraphIndex = findParagraphWithPattern(body, pattern, startIndex)
    if (paragraphIndex === null) {
        throw new Error(`The template document should have pattern ${pattern}`)
    }

    let text = getParagraphText(body[paragraphIndex])
    if (text != pattern) {
        throw new Error(`The ${pattern} pattern should be the only text of the paragraph`)
    }

    return paragraphIndex
}

function getDocumentBody(document: any): any {
    let documentTag = XML.getChildTag(document, "w:document")["w:document"]
    return XML.getChildTag(documentTag, "w:body")["w:body"]
}

function getMetaString(value: any): string {
    if (Array.isArray(value)) {
        let result = ""
        for (let component of value) {
            result += getMetaString(component)
        }
        return result
    }

    if (typeof value !== "object" || !value.t) {
        return ""
    }

    if (value.t === "Str") {
        return value.c
    }
    if (value.t === "Strong") {
        return "__" + getMetaString(value.c) + "__"
    }
    if (value.t === "Emph") {
        return "_" + getMetaString(value.c) + "_"
    }
    if (value.t === "Cite") {
        return getMetaString(value.c[1])
    }
    if (value.t === "Space") {
        return " "
    }
    if (value.t === "Link") {
        return getMetaString(value.c[1])
    }

    return getMetaString(value.c)
}

function convertMetaToJsonRecursive(meta: any): any{
    if (meta.t === "MetaList") {
        return meta.c.map((element) => {
            return convertMetaToJsonRecursive(element)
        })
    }

    if (meta.t === "MetaMap") {
        let result = {}
        for (let key of Object.getOwnPropertyNames(meta.c)) {
            result[key] = convertMetaToJsonRecursive(meta.c[key])
        }
        return result
    }

    if (meta.t === "MetaInlines") {
        return getMetaString(meta.c)
    }
}

function convertMetaToObject(meta: any): any {
    let result = {}
    for (let key of Object.getOwnPropertyNames(meta)) {
        result[key] = convertMetaToJsonRecursive(meta[key])
    }
    return result
}

function templateReplaceBodyContents(templateBody: any, body: any) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{body}}}")

    templateBody.splice(paragraphIndex, 1, ...body)
}

function clearParagraphContents(paragraph: any): any {
    let contents = paragraph["w:p"]

    for (let i = 0; i < contents.length; i++) {
        let tagName = XML.getTagName(contents[i])
        if (tagName === "w:r") {
            contents.splice(i, 1)
            i--
        }
    }
}

function templateAuthorList(templateBody: any, meta: any) {

    let authors = meta["ispras_templates"].authors

    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{authors_${language}}}}`)

        let newParagraphs = []

        let authorIndex = 1;

        for (let author of authors) {
            let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]))
            clearParagraphContents(newParagraph)

            let indexLine = String(authorIndex)
            let authorLine = author["name_" + language] + ", ORCID: " + author.orcid + ", <" + author.email + ">"

            let indexTag = OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()])
            let authorTag = OXML.buildParagraphTextTag(authorLine)

            newParagraph["w:p"].push(indexTag, authorTag)
            newParagraphs.push(newParagraph)

            authorIndex++
        }

        templateBody.splice(paragraphIndex, 1, ...newParagraphs)
    }

    for (let language of languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{organizations_${language}}}}`)
        let organizations = meta["ispras_templates"]["organizations_" + language]

        let newParagraphs = []
        let orgIndex = 1

        for (let organizationLine of organizations) {
            let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]))
            clearParagraphContents(newParagraph)

            let indexLine = String(orgIndex)

            let indexTag = OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()])
            let organizationTag = OXML.buildParagraphTextTag(organizationLine)

            newParagraph["w:p"].push(indexTag, organizationTag)
            newParagraphs.push(newParagraph)

            orgIndex++
        }

        templateBody.splice(paragraphIndex, 1, ...newParagraphs)
    }
}

function templateReplaceLinks(templateBody: any, meta: any, listRules: any) {
    let litListRule = listRules["LitList"]
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{links}}}")
    let links = meta["ispras_templates"].links

    let newParagraphs = []

    for (let link of links) {
        let newParagraph = OXML.buildParagraphWithStyle(litListRule.styleName)
        let style = XML.getChildTag(newParagraph["w:p"], "w:pPr")
        style["w:pPr"].push(OXML.buildNumPr("0", litListRule.numId))

        newParagraph["w:p"].push(OXML.buildParagraphTextTag(link))
        newParagraphs.push(newParagraph)
    }

    templateBody.splice(paragraphIndex, 1, ...newParagraphs)
}

function templateReplaceAuthorsDetail(templateBody: any, meta: any) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{authors_detail}}}")
    let authors = meta["ispras_templates"].authors

    let newParagraphs = []

    for (let author of authors) {
        for (let language of languages) {
            let newParagraph = JSON.parse(JSON.stringify(templateBody[paragraphIndex]))

            let line = author["details_" + language]

            clearParagraphContents(newParagraph)
            newParagraph["w:p"].push(OXML.buildParagraphTextTag(line))
            newParagraphs.push(newParagraph)
        }
    }

    templateBody.splice(paragraphIndex, 1, ...newParagraphs)
}

function replacePageHeaders(headers: any[], meta: any): any {
    let header_ru = meta["ispras_templates"].page_header_ru
    let header_en = meta["ispras_templates"].page_header_en

    if (header_ru === "@use_citation") {
        header_ru = meta["ispras_templates"].for_citation_ru
    }

    if (header_en === "@use_citation") {
        header_en = meta["ispras_templates"].for_citation_en
    }

    for (let header of headers) {
        replaceInlineTemplate(header, `{{{page_header_ru}}}`, header_ru)
        replaceInlineTemplate(header, `{{{page_header_en}}}`, header_en)
    }
}

function replaceTemplates(template: any, body: any, meta: any): any {
    let templateCopy = JSON.parse(JSON.stringify(template))

    let templateBody = getDocumentBody(templateCopy)

    templateReplaceBodyContents(templateBody, body)
    templateAuthorList(templateBody, meta)

    let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"]

    for (let template of templates) {
        for (let language of languages) {
            let template_lang = template + "_" + language
            let value = meta["ispras_templates"][template_lang]
            replaceInlineTemplate(templateBody, `{{{${template_lang}}}}`, value)
        }
    }

    templateReplaceAuthorsDetail(templateBody, meta)

    return templateCopy
}

function setXmlns(xml: any, xmlns: Map<string, string>) {
    let documentTag = XML.getChildTag(xml, "w:document")

    for (let [key, value] of xmlns) {
        documentTag[XML.keys.attributes][key] = value
    }
}

function patchRelIds(doc: any, map: Map<string, string>) {
    if (Array.isArray(doc)) {
        for (let child of doc) {
            patchRelIds(child, map)
        }
    }

    if (typeof doc != "object") return

    let tagName = XML.getTagName(doc)

    let attrs = doc[XML.keys.attributes]
    if (attrs) {
        for (let attr in ["r:id", "r:embed"]) {
            let relId = attrs[attr]
            if (relId && map.has(relId)) {
                attrs[attr] = map.get(relId)
            }
        }
    }

    if (doc[XML.keys.attributes]) {
        let relIdAttr = tagsWithRelId.get(tagName)
        if (relIdAttr) {
            let relId = doc[XML.keys.attributes][relIdAttr]
            if (relId && map.has(relId)) {
                doc[XML.keys.attributes][relIdAttr] = map.get(relId)
            }
        }
    }

    patchRelIds(doc[tagName], map)
}

async function fixDocxStyles(sourcePath, targetPath, meta): Promise<void> {
    let resourcesDir = path.dirname(process.argv[1]) + "/../resources"

    // Load the source and target documents
    let target = await JSZip.loadAsync(fs.readFileSync(sourcePath))
    let source = await JSZip.loadAsync(fs.readFileSync(resourcesDir + '/isp-reference.docx'))

    let sourceStylesXML = await source.file("word/styles.xml").async("string");
    let targetStylesXML = await target.file("word/styles.xml").async("string");
    let sourceDocXML = await source.file("word/document.xml").async("string");
    let targetDocXML = await target.file("word/document.xml").async("string");
    let targetContentTypesXML = await target.file("[Content_Types].xml").async("string");
    let targetDocumentRelsXML = await target.file("word/_rels/document.xml.rels").async("string");
    let sourceDocumentRelsXML = await source.file("word/_rels/document.xml.rels").async("string");
    let targetNumberingXML = await source.file("word/numbering.xml").async("string");
    let sourceHeader1 = await source.file("word/header1.xml").async("string");
    let sourceHeader2 = await source.file("word/header2.xml").async("string");
    let sourceHeader3 = await source.file("word/header3.xml").async("string");

    let targetContentTypesParsed = XML.parser.parse(targetContentTypesXML);
    let targetDocumentRelsParsed = XML.parser.parse(targetDocumentRelsXML);
    let sourceDocumentRelsParsed = XML.parser.parse(sourceDocumentRelsXML);
    let sourceStylesParsed = XML.parser.parse(sourceStylesXML);
    let targetStylesParsed = XML.parser.parse(targetStylesXML);
    let sourceDocParsed = XML.parser.parse(sourceDocXML);
    let targetDocParsed = XML.parser.parse(targetDocXML);
    let targetNumberingParsed = XML.parser.parse(targetNumberingXML);
    let sourceHeader1Parsed = XML.parser.parse(sourceHeader1)
    let sourceHeader2Parsed = XML.parser.parse(sourceHeader2)
    let sourceHeader3Parsed = XML.parser.parse(sourceHeader3)

    copyLatentStyles(sourceStylesParsed, targetStylesParsed)
    copyDocDefaults(sourceStylesParsed, targetStylesParsed)

    let targetStylesNamesToId = getStyleIdsByNameFromDefs(XML.getChildTag(targetStylesParsed, "w:styles")["w:styles"]);
    let sourceStylesNamesToId = getStyleIdsByNameFromDefs(XML.getChildTag(sourceStylesParsed, "w:styles")["w:styles"]);

    let sourceStyleTable = getStyleTable(sourceStylesParsed)

    let usedStyles = getUsedStylesDeep(sourceDocParsed, sourceStyleTable, [
        "ispSubHeader-1 level",
        "ispSubHeader-2 level",
        "ispSubHeader-3 level",
        "ispAuthor",
        "ispAnotation",
        "ispText_main",
        "ispList",
        "ispListing",
        "ispListing Знак",
        "ispLitList",
        "ispPicture_sign",
        "ispNumList",
        "Normal"
    ].map(name => sourceStylesNamesToId.get(name)))
    let mappingTable = getMappingTable(usedStyles)

    patchStyleDefinitions(sourceDocParsed, sourceStylesParsed, mappingTable)
    patchStyleUseReferences(sourceDocParsed, sourceStylesParsed, mappingTable)
    let extractedDefs = extractStyleDefs(sourceStylesParsed)
    let extractedStyleIdsByName = getStyleIdsByNameFromDefs(extractedDefs)

    let stylePatch = new Map<string, string>([
        ["Heading1", extractedStyleIdsByName.get("ispSubHeader-1 level")],
        ["Heading2", extractedStyleIdsByName.get("ispSubHeader-2 level")],
        ["Heading3", extractedStyleIdsByName.get("ispSubHeader-3 level")],
        ["Author", extractedStyleIdsByName.get("ispAuthor")],
        ["AbstractTitle", extractedStyleIdsByName.get("ispAnotation")],
        ["Abstract", extractedStyleIdsByName.get("ispAnotation")],
        ["BlockText", extractedStyleIdsByName.get("ispText_main")],
        ["BodyText", extractedStyleIdsByName.get("ispText_main")],
        ["FirstParagraph", extractedStyleIdsByName.get("ispText_main")],
        ["Normal", extractedStyleIdsByName.get("Normal")],
        ["SourceCode", extractedStyleIdsByName.get("ispListing")],
        ["VerbatimChar", extractedStyleIdsByName.get("ispListing Знак")],
        ["ImageCaption", extractedStyleIdsByName.get("ispPicture_sign")],
    ])

    let stylesToRemove = new Set<string>([
        "Heading4",
        "Heading5",
        "Heading6",
        "Heading7",
        "Heading8",
        "Heading9",
    ])

    for (let possibleCollision of extractedStyleIdsByName) {
        let sourceStyleName = possibleCollision[0]
        let sourceStyleId = possibleCollision[1]

        if (targetStylesNamesToId.has(sourceStyleName)) {
            let targetStyleId = targetStylesNamesToId.get(sourceStyleName)

            if (!stylePatch.has(targetStyleId)) {
                stylePatch.set(targetStyleId, sourceStyleId)
            }
            stylesToRemove.add(targetStyleId)
        }
    }

    removeCollidedStyles(targetStylesParsed, stylesToRemove)

    appendStyles(targetStylesParsed, extractedDefs)

    patchStyleUseReferences(targetDocParsed, targetStylesParsed, stylePatch)

    let patchRules = {
        "OrderedList": {styleName: extractedStyleIdsByName.get("ispNumList"), numId: "33"},
        "BulletList": {styleName: extractedStyleIdsByName.get("ispList1"), numId: "43"},
        "LitList": {styleName: extractedStyleIdsByName.get("ispLitList"), numId: "80"},
    };

    let newListStyles = applyListStyles(targetDocParsed, patchRules)

    setXmlns(sourceDocParsed, properDocXmlns)

    let relMap = transferRels(sourceDocumentRelsParsed, targetDocumentRelsParsed)
    patchRelIds(sourceDocParsed, relMap)

    targetDocParsed = replaceTemplates(sourceDocParsed, getDocumentBody(targetDocParsed), meta)

    templateReplaceLinks(getDocumentBody(targetDocParsed), meta, patchRules)

    addNewNumberings(targetNumberingParsed, newListStyles)

    replacePageHeaders([sourceHeader1Parsed, sourceHeader2Parsed, sourceHeader3Parsed], meta)

    addContentType(targetContentTypesParsed, "/word/footer1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
    addContentType(targetContentTypesParsed, "/word/footer2.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
    addContentType(targetContentTypesParsed, "/word/footer3.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
    addContentType(targetContentTypesParsed, "/word/header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
    addContentType(targetContentTypesParsed, "/word/header2.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
    addContentType(targetContentTypesParsed, "/word/header3.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")

    await copyFile(source, target, "word/_rels/header1.xml.rels")
    await copyFile(source, target, "word/_rels/header2.xml.rels")
    await copyFile(source, target, "word/_rels/header3.xml.rels")
    await copyFile(source, target, "word/_rels/footer1.xml.rels")
    await copyFile(source, target, "word/_rels/footer2.xml.rels")
    await copyFile(source, target, "word/_rels/footer3.xml.rels")

    await copyFile(source, target, "word/footer1.xml")
    await copyFile(source, target, "word/footer2.xml")
    await copyFile(source, target, "word/footer3.xml")

    await copyFile(source, target, "word/footnotes.xml")
    await copyFile(source, target, "word/theme/theme1.xml")
    await copyFile(source, target, "word/fontTable.xml")
    await copyFile(source, target, "word/settings.xml")
    await copyFile(source, target, "word/webSettings.xml")
    await copyFile(source, target, "word/media/image1.png")

    target.file("word/header1.xml", XML.builder.build(sourceHeader1Parsed))
    target.file("word/header2.xml", XML.builder.build(sourceHeader2Parsed))
    target.file("word/header3.xml", XML.builder.build(sourceHeader3Parsed))

    target.file("word/_rels/document.xml.rels", XML.builder.build(targetDocumentRelsParsed))
    target.file("[Content_Types].xml", XML.builder.build(targetContentTypesParsed))
    target.file("word/numbering.xml", XML.builder.build(targetNumberingParsed))
    target.file("word/styles.xml", XML.builder.build(targetStylesParsed))
    target.file("word/document.xml", XML.builder.build(targetDocParsed))

    fs.writeFileSync(targetPath, await target.generateAsync({type: "uint8array"}));
}

function fixCompactLists(list): any[] {
    // For compact list, 'para' is replaced with 'plain'.
    // Compact lists were not mentioned in the
    // guidelines, so get rid of them

    for (let i = 0; i < list.c.length; i++) {
        let element = list.c[i]
        if (typeof element[0] === "object" && element[0].t === "Plain") {
            element[0].t = "Para"
        }
        list.c[i] = getPatchedMetaElement(list.c[i])
    }

    return [
        {
            t: "RawBlock",
            c: ["openxml", `<!-- ListMode ${list.t} -->`]
        },
        list,
        {
            t: "RawBlock",
            c: ["openxml", `<!-- ListMode None -->`]
        }
    ]
}

function getImageCaption(content): any {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ...XML.buildAttributes({"w:val": "ImageCaption"})
                }, {
                    "w:contextualSpacing": [],
                    ...XML.buildAttributes({"w:val": "true"})
                }]
        },
        OXML.buildParagraphTextTag(getMetaString(content))
    ];

    return {
        t: "RawBlock",
        c: ["openxml", `<w:p>${XML.builder.build(elements)}</w:p>`]
    };
}

function getListingCaption(content): any {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ...XML.buildAttributes({"w:val": "BodyText"})
                }, {
                    "w:jc": [],
                    ...XML.buildAttributes({"w:val": "left"})
                },
            ]
        },
        OXML.buildParagraphTextTag(getMetaString(content), [
            {"w:i": []},
            {"w:iCs": []},
            {"w:sz": [], ...XML.buildAttributes({"w:val": "18"})},
            {"w:szCs": [], ...XML.buildAttributes({"w:val": "18"})},
        ])
    ];

    return {
        t: "RawBlock",
        c: ["openxml", `<w:p>${XML.builder.build(elements)}</w:p>`]
    };
}

function getPatchedMetaElement(element): any {
    if (Array.isArray(element)) {
        let newArray = []

        for (let i = 0; i < element.length; i++) {
            let patched = getPatchedMetaElement(element[i])
            if (Array.isArray(patched) && !Array.isArray(element[i])) {
                newArray.push(...patched)
            } else {
                newArray.push(patched)
            }
        }

        return newArray
    }

    if (typeof element !== "object" || !element) {
        return element
    }

    let type = element.t
    let value = element.c

    if (type === 'Div') {
        let content = value[1];
        let classes = value[0][1];
        if (classes) {
            if (classes.includes("img-caption")) {
                return getImageCaption(content)
            }

            if (classes.includes("table-caption") || classes.includes("listing-caption")) {
                return getListingCaption(content)
            }
        }
    } else if (type === 'BulletList' || type === 'OrderedList') {
        return fixCompactLists(element)
    }

    for (let key of Object.getOwnPropertyNames(element)) {
        element[key] = getPatchedMetaElement(element[key]);
    }

    return element
}

async function generatePandocDocx(source: string, target: string): Promise<any> {
    let markdown = await fs.promises.readFile(source, "utf-8")

    let meta = await pandoc(markdown, ["-f", "markdown", "-t", "json", ...pandocFlags])
    let metaParsed = JSON.parse(meta)

    metaParsed.blocks = getPatchedMetaElement(metaParsed.blocks)

    await pandoc(JSON.stringify(metaParsed), ["-f", "json", "-t", "docx", "-o", target])

    return convertMetaToObject(metaParsed.meta)
}

async function main(): Promise<void> {
    let argv = process.argv
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>")
        process.exit(1)
    }

    let source = argv[2]
    let target = argv[3]

    let meta = await generatePandocDocx(source, target + ".tmp")
    await fixDocxStyles(target + ".tmp", target, meta).then()
}

main().then()