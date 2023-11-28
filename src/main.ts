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
import PandocJsonPatcher from "src/pandoc/pandoc-json-patcher";
import {PandocJsonMeta} from "src/pandoc/pandoc-json-meta";
import {PandocJson} from "src/pandoc/pandoc-json";

const pandocFlags = ["--tab-stop=8"]
export const languages = ["ru", "en"]
const resourcesDir = path.dirname(process.argv[1]) + "/../resources"

function getLinksParagraphs(document: WordDocument, meta: PandocJsonMeta) {
    let styleId = document.styles.resource.getStyleByName("ispLitList").getId()
    let numId = "80"
    let links = meta.getSection("links").asArray()

    let result = []

    for (let link of links) {
        let paragraph = OXML.buildParagraphWithStyle(styleId)
        let style = paragraph.getChild("w:pPr")
        style.pushChild(OXML.buildNumPr("0", numId))

        paragraph.pushChild(OXML.buildParagraphTextTag(link.getString()))
        result.push(paragraph)
    }

    return result
}

function getAuthors(document: WordDocument, meta: PandocJsonMeta, language: string) {
    let styleId = document.styles.resource.getStyleByName("ispAuthor").getId()
    let authors = meta.getSection("authors").asArray()

    let result = []
    let authorIndex = 1;

    for (let author of authors) {
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        let name = author.getString("name_" + language)
        let orcid = author.getString("orcid")
        let email = author.getString("email")

        let indexLine = String(authorIndex)
        let authorLine = `${name}, ORCID: ${orcid}, <${email}>`

        paragraph.pushChild(OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()]))
        paragraph.pushChild(OXML.buildParagraphTextTag(authorLine))

        result.push(paragraph)

        authorIndex++
    }

    return result
}

function getOrganizations(document: WordDocument, meta: PandocJsonMeta, language: string) {
    let styleId = document.styles.resource.getStyleByName("ispAuthor").getId()
    let organizations = meta.getSection("organizations_" + language).asArray()

    let orgIndex = 1
    let result = []

    for (let organization of organizations) {
        let paragraph = OXML.buildParagraphWithStyle(styleId)

        let indexLine = String(orgIndex)

        paragraph.pushChild(OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()]))
        paragraph.pushChild(OXML.buildParagraphTextTag(organization.getString()))

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

function getImageCaption(document: WordDocument, content: string): XML.Node {
    // This function is called from patchPandocJson, so this caption is inserted in
    // the content document, not in the template document.
    // "Image Caption" is a pandoc style that later gets converted to "ispPicture_sign"

    let styleId = document.styles.resource.getStyleByName("Image Caption").getId()

    return XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", styleId),
            XML.Node.build("w:contextualSpacing").setAttr("w:val", "true"),
        ]),
        OXML.buildParagraphTextTag(content)
    ]);
}

function getListingCaption(document: WordDocument, content: string): XML.Node {
    // Same note here:
    // "Body Text" is a pandoc style that later gets converted to "ispText_main"

    let styleId = document.styles.resource.getStyleByName("Body Text").getId()

    return XML.Node.build("w:p").appendChildren([
        XML.Node.build("w:pPr").appendChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", styleId),
            XML.Node.build("w:jc").setAttr("w:val", "left"),
        ]),
        OXML.buildParagraphTextTag(content, [
            XML.Node.build("w:i"),
            XML.Node.build("w:iCs"),
            XML.Node.build("w:sz").setAttr("w:val", "18"),
            XML.Node.build("w:szCs").setAttr("w:val", "18"),
        ])
    ])
}

function patchPandocJson(contentDoc: WordDocument, pandocJson: PandocJson) {
    new PandocJsonPatcher(pandocJson)
        .replaceDivWithClass("img-caption", (contents) => getImageCaption(contentDoc, contents))
        .replaceDivWithClass("table-caption", (contents) => getListingCaption(contentDoc, contents))
        .replaceDivWithClass("listing-caption", (contents) => getListingCaption(contentDoc, contents))
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
    let contentDoc = await new WordDocument().load(tmpDocPath)
    let markdown = await fs.promises.readFile(markdownSource, "utf-8")
    let pandocJson = await pandoc.markdownToPandocJson(markdown, pandocFlags)

    patchPandocJson(contentDoc, pandocJson)

    await pandoc.pandocJsonToDocx(pandocJson, ["-o", tmpDocPath])
    let pandocJsonMeta = new PandocJsonMeta(pandocJson.meta["ispras_templates"])
    let templateDoc = await new WordDocument().load(resourcesDir + '/isp-reference.docx')

    await patchTemplateDocx(templateDoc, contentDoc, pandocJsonMeta)

    await templateDoc.save(targetPath)
}

main().then()