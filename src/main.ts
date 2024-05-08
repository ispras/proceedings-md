import * as path from 'path';
import * as fs from 'fs';
import * as pandoc from "src/pandoc/pandoc";
import * as XML from "src/xml";
import * as OXML from "src/word/oxml";
import * as process from "process";
import WordDocument from "src/word/word-document";
import {StyledTemplateSubstitution} from "src/word-templates/styled-template-substitution";
import InlineTemplateSubstitution from "src/word-templates/inline-template-substitution";
import ParagraphTemplateSubstitution from "src/word-templates/paragraph-template-substitution";
import PandocJsonPatcher, {getOpenxmlInjection} from "src/pandoc/pandoc-json-patcher";
import {PandocJsonMeta} from "src/pandoc/pandoc-json-meta";
import {MetaElement, PandocJson} from "src/pandoc/pandoc-json";

const pandocFlags = ["--tab-stop=8"]
export const languages = ["ru", "en"]
const resourcesDir = path.dirname(process.argv[1]) + "/../resources"

function getLinksParagraphs(document: WordDocument, meta: PandocJsonMeta) {
    let styleId = document.styles.resource.getStyleByName("ispLitList").getId()
    let numId = "80"
    let linksSection = meta.getSection("links").asArray()

    let result = []

    for (let link of linksSection) {
        let description: string
        // Backwards compatibility
        if (link.isMap()) {
            description = link.getString("description")
        } else {
            description = link.getString()
        }
        let paragraph = OXML.buildParagraphWithStyle(styleId)
        let style = paragraph.getChild("w:pPr")
        style.pushChild(OXML.buildNumPr("0", numId))

        paragraph.pushChild(OXML.buildParagraphTextTag(description))
        result.push(paragraph)
    }

    return result
}

function getAuthors(document: WordDocument, meta: PandocJsonMeta, language: string) {
    let styleId = document.styles.resource.getStyleByName("ispAuthor").getId()
    let authors = meta.getSection("authors").asArray()
    let organizations = meta
        .getSection("organizations")
        .asArray()
        .map(section => section.getString("id"))

    let result = []

    for (let author of authors) {
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        let name = author.getString("name_" + language)
        let orcid = author.getString("orcid")
        let email = author.getString("email")
        let authorOrgs = author.getSection("organizations")
            .asArray()
            .map(section => section.getString())
            .map(id => organizations.indexOf(id) + 1)
            .join(",")

        let authorLine = `${name}, ORCID: ${orcid}, <${email}>`

        paragraph.pushChild(OXML.buildParagraphTextTag(authorOrgs, [OXML.buildSuperscriptTextStyle()]))
        paragraph.pushChild(OXML.buildParagraphTextTag(authorLine))

        result.push(paragraph)
    }

    return result
}

function getOrganizations(document: WordDocument, meta: PandocJsonMeta, language: string) {
    let styleId = document.styles.resource.getStyleByName("ispAuthor").getId()
    let organizations = meta.getSection("organizations").asArray()

    let orgIndex = 1
    let result = []

    for (let organization of organizations) {
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        let indexLine = String(orgIndex)

        paragraph.pushChild(OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()]))
        paragraph.pushChild(OXML.buildParagraphTextTag(organization.getString("name_" + language)))

        result.push(paragraph)

        orgIndex++
    }

    return result
}

function getAuthorsDetail(document: WordDocument, meta: PandocJsonMeta) {
    let styleId = document.styles.resource.getStyleByName("ispText_main").getId()
    let authors = meta.getSection("authors").asArray()

    let result = []

    for (let author of authors) {
        for (let language of languages) {
            let line = author.getString("details_" + language)
            let newParagraph = OXML.buildParagraphWithStyle(styleId)
            newParagraph.getChild("w:pPr").pushChild(
                XML.Node.build("w:spacing")
                    .setAttr("w:before", "30")
                    .setAttr("w:after", "120")
            )
            newParagraph.pushChild(OXML.buildParagraphTextTag(line))
            result.push(newParagraph)
        }
    }

    return result
}

function getImageCaption(content: string) {
    // This function is called from patchPandocJson, so this caption is inserted in
    // the content document, not in the template document.
    // "Image Caption" is a pandoc style that gets converted to "ispPicture_sign" later
    // Unfortunately, style table is not available yet, but it seems that Image Caption style consistently gets converted to
    // "ImageCaption". It's yet to be asserted.

    // let styleId = document.styles.resource.getStyleByName("Image Caption").getId()
    let styleId = "ImageCaption"

    let node = XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", styleId),
            XML.Node.build("w:contextualSpacing").setAttr("w:val", "true"),
        ]),
        OXML.buildParagraphTextTag(content)
    ]);

    return getOpenxmlInjection(node)
}

function getListingCaption(content: string) {
    // Same note here:
    // "Body Text" is a pandoc style that gets converted to "ispText_main" later
    // Unfortunately, style table is not available yet, but it seems that Body Text style consistently gets converted to
    // "BodyText"

    // let styleId = document.styles.resource.getStyleByName("Body Text").getId()
    let styleId = "BodyText"

    let node = XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", styleId),
            XML.Node.build("w:jc").setAttr("w:val", "left")
        ]),
        OXML.buildParagraphTextTag(content, [
            XML.Node.build("w:i"),
            XML.Node.build("w:iCs"),
            XML.Node.build("w:sz").setAttr("w:val", "18"),
            XML.Node.build("w:szCs").setAttr("w:val", "18"),
        ])
    ])

    return getOpenxmlInjection(node)
}

function checkStyleIds(document: WordDocument) {

    function check(style: string, expected: string) {
        let bodyTextId = document.styles.resource.getStyleByName(style).getId()
        if (bodyTextId !== expected) {
            console.warn("Your pandoc version has 'Body Text' style with id '" + bodyTextId + "' instead of '" + expected + "'. Some text styles can be corrupted.")
        }
    }

    check("Body Text", "BodyText")
    check("Image Caption", "ImageCaption")
}

class DocumentReferences {
    stack: number[] = []
    depthThreshold: number = 1
    groups = new Map<string, Map<string, string>>()
    meta: PandocJsonMeta

    constructor(meta: PandocJsonMeta) {
        this.meta = meta
    }

    getPrefixMap(prefix: string) {
        let map = this.groups.get(prefix)
        if (!map) {
            map = new Map()
            this.groups.set(prefix, map)
        }
        return map
    }

    getSection(header: MetaElement<"Header">) {
        let depth = Math.max(0, header.c[0] - this.depthThreshold)
        let label = header.c[1][0]

        let prefix = label.split(":")[0]
        let map = this.getPrefixMap(prefix)

        if (map.has(label)) {
            console.warn("Multiple definitions of section " + label)
        }

        while (this.stack.length > depth) this.stack.pop()
        if (this.stack.length === depth) {
            this.stack[this.stack.length - 1]++
        } else {
            while (this.stack.length < depth) this.stack.push(1)
        }

        if(this.stack.length === 0) {
            return
        }

        let result = this.stack.join(".")
        map.set(label, result)

        if(this.stack.length == 1) {
            result += "."
        }

        header.c[2].unshift({
            "t": "Str",
            "c": result
        }, {
            "t": "Space",
            "c": undefined
        })
    }

    getReference(reference: string): MetaElement<"Str"> {
        let prefix = reference.split(":")[0]
        let map = this.getPrefixMap(prefix)

        let index = map.get(reference)
        if (index === undefined) {
            index = (map.size + 1).toString()
            map.set(reference, index)
        }

        return {
            "t": "Str",
            "c": index.toString()
        }
    }

    getCite(reference: string): MetaElement<"Str"> {
        let links = this.meta.getSection("links").asArray()
        for (let i = 0; i < links.length; i++) {
            let link = links[i]
            if (link.isMap() && link.getString("id") === reference) {
                return {
                    "t": "Str",
                    "c": "[" + (i + 1).toString() + "]"
                }
            }
        }

        console.warn("Undefined citation: " + reference)

        return {
            "t": "Str",
            "c": "[?]"
        }
    }
}

function patchPandocJson(pandocJson: PandocJson, meta: PandocJsonMeta) {
    let references = new DocumentReferences(meta)

    new PandocJsonPatcher(pandocJson)
        .replaceElements("Header", (contents) => references.getSection(contents))
        .replaceSpanWithClass("cite", (contents) => references.getCite(contents))
        .replaceSpanWithClass("ref", (contents) => references.getReference(contents))
        .replaceDivWithClass("img-caption", (contents) => getImageCaption(contents))
        .replaceDivWithClass("table-caption", (contents) => getListingCaption(contents))
        .replaceDivWithClass("listing-caption", (contents) => getListingCaption(contents))
}

function centerDrawings(doc: WordDocument) {
    let document = doc.document.resource.toXml()
    document.visitSubtree("w:drawing", (node: XML.Node, path: XML.Path) => {
        let parentPath = path.slice(0, -2)
        let parent = document.getChild(parentPath)
        if (parent.getTagName() !== "w:p") return

        parent.getChild("w:pPr").appendChildren([
            XML.Node.build("w:jc").setAttr("w:val", "center")
        ])
    })
}

async function patchTemplateDocx(templateDoc: WordDocument, contentDoc: WordDocument, pandocJsonMeta: PandocJsonMeta) {
    await new StyledTemplateSubstitution()
        .setSource(contentDoc)
        .setTarget(templateDoc)
        .setTemplate("{{{body}}}")
        .setStyleConversion(new Map([
            ["Heading 1", "ispSubHeader-1 level"],
            ["Heading 2", "ispSubHeader-2 level"],
            ["Heading 3", "ispSubHeader-3 level"],
            ["Heading 4", "ispSubHeader-3 level"],
            ["Author", "ispAuthor"],
            ["Abstract Title", "ispAnotation"],
            ["Abstract", "ispAnotation"],
            ["Block Text", "ispText_main"],
            ["Body Text", "ispText_main"],
            ["First Paragraph", "ispText_main"],
            ["Normal", "Normal"],
            ["Compact", "Normal"],
            ["Source Code", "ispListing"],
            ["Verbatim Char", "ispListing Знак"],
            ["Image Caption", "ispPicture_sign"],
            ["Table", "Table Grid"]
        ]))
        .setStylesToMigrate(new Set([
            ...pandoc.tokenClasses
        ]))
        .setAllowUnrecognizedStyles(false)
        .setListConversion({
            decimal: {
                styleName: "ispNumList",
                numId: "33"
            },
            bullet: {
                styleName: "ispList1",
                numId: "43"
            }
        })
        .perform();

    let inlineSubstitution = new InlineTemplateSubstitution().setDocument(templateDoc)
    let paragraphSubstitution = new ParagraphTemplateSubstitution().setDocument(templateDoc)

    for (let language of languages) {
        let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"]
        for (let template of templates) {

            let template_lang = template + "_" + language
            let replacement = pandocJsonMeta.getString(template_lang)

            inlineSubstitution
                .setTemplate("{{{" + template_lang + "}}}")
                .setReplacement(replacement)
                .perform()
        }

        let header = pandocJsonMeta.getString("page_header_" + language)

        if (header === "@use_citation") {
            header = pandocJsonMeta.getString("for_citation_" + language)
        }

        inlineSubstitution
            .setTemplate("{{{page_header_" + language + "}}}")
            .setReplacement(header)
            .perform()

        paragraphSubstitution
            .setTemplate("{{{authors_" + language + "}}}")
            .setReplacement(() => getAuthors(templateDoc, pandocJsonMeta, language))
            .perform()

        paragraphSubstitution
            .setTemplate("{{{organizations_" + language + "}}}")
            .setReplacement(() => getOrganizations(templateDoc, pandocJsonMeta, language))
            .perform()
    }

    paragraphSubstitution
        .setTemplate("{{{links}}}")
        .setReplacement(() => getLinksParagraphs(templateDoc, pandocJsonMeta))
        .perform()

    paragraphSubstitution
        .setTemplate("{{{authors_detail}}}")
        .setReplacement(() => getAuthorsDetail(templateDoc, pandocJsonMeta))
        .perform()

    centerDrawings(templateDoc)
}

async function main(): Promise<void> {
    let argv = process.argv
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>")
        process.exit(1)
    }

    let markdownSource = argv[2]
    let targetPath = argv[3]

    let tmpDocPath = targetPath + ".tmp"
    let markdown = await fs.promises.readFile(markdownSource, "utf-8")
    let pandocJson = await pandoc.markdownToPandocJson(markdown, pandocFlags)
    let pandocJsonMeta = new PandocJsonMeta(pandocJson.meta["ispras_templates"])

    await fs.promises.writeFile(markdownSource + ".json", JSON.stringify(pandocJson, null, 4), "utf-8")
    patchPandocJson(pandocJson, pandocJsonMeta)
    await fs.promises.writeFile(markdownSource + ".patched.json", JSON.stringify(pandocJson, null, 4), "utf-8")

    await pandoc.pandocJsonToDocx(pandocJson, ["-o", tmpDocPath])

    let templateDoc = await new WordDocument().load(resourcesDir + '/isp-reference.docx')

    let contentDoc = await new WordDocument().load(tmpDocPath)
    checkStyleIds(contentDoc)
    await patchTemplateDocx(templateDoc, contentDoc, pandocJsonMeta)

    await templateDoc.save(targetPath)
}

main().then()