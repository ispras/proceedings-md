"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.languages = void 0;
const path = __importStar(require("path"));
const fs = __importStar(require("fs"));
const JSZip = __importStar(require("jszip"));
const pandoc = __importStar(require("./pandoc"));
const XML = __importStar(require("./xml"));
const OXML = __importStar(require("./oxml"));
const pandoc_1 = require("./pandoc");
const pandocFlags = ["--tab-stop=8"];
const properDocXmlns = new Map([
    ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
    ["xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math"],
    ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["xmlns:o", "urn:schemas-microsoft-com:office:office"],
    ["xmlns:v", "urn:schemas-microsoft-com:vml"],
    ["xmlns:w10", "urn:schemas-microsoft-com:office:word"],
    ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
    ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
    ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
]);
exports.languages = ["ru", "en"];
function visitStyleCrossReferences(style, callback) {
    let basedOnTag = style.getChild("w:basedOn");
    if (basedOnTag)
        callback(basedOnTag);
    let linkTag = style.getChild("w:link");
    if (linkTag)
        callback(linkTag);
    let nextTag = style.getChild("w:next");
    if (nextTag)
        callback(nextTag);
}
function getStyleCrossReferences(styles) {
    let result = [];
    styles.getChild("w:styles").visitChildren("w:style", (style) => {
        result.push(style.shallowCopy());
        visitStyleCrossReferences(style, (node) => result.push(node));
    });
    return result;
}
function getDocStyleUseReferences(doc, result = []) {
    doc.visitSubtree((node) => {
        return node.getTagName() === "w:pStyle" || node.getTagName() == "w:rStyle";
    }, (node) => {
        result.push(node.shallowCopy());
    });
    return result;
}
function extractStyleDefs(styles) {
    let result = [];
    styles.getChild("w:styles").visitChildren("w:style", (style) => {
        result.push(style.deepCopy());
    });
    return result;
}
function extractTemplateStyleDefs(styles) {
    let result = [];
    styles.getChild("w:styles").visitChildren("w:style", (style) => {
        if (style.getAttr("w:styleId").startsWith("template-")) {
            result.push(style.deepCopy());
        }
    });
    return result;
}
function patchStyleDefinitions(doc, styles, map) {
    styles.getChild("w:styles").visitChildren("w:style", (style) => {
        if (style.getAttr("w:styleId") && map.has(style.getAttr("w:styleId"))) {
            style.setAttr("w:styleId", map.get(style.getAttr("w:styleId")));
        }
    });
}
function patchStyleUseReferences(doc, styles, map) {
    let docReferences = getDocStyleUseReferences(doc);
    let crossReferences = getStyleCrossReferences(styles);
    for (let ref of docReferences.concat(crossReferences)) {
        if (ref.getAttr("w:val") && map.has(ref.getAttr("w:val"))) {
            ref.setAttr("w:val", map.get(ref.getAttr("w:val")));
        }
    }
}
function getUsedStyles(doc) {
    let references = getDocStyleUseReferences(doc);
    let set = new Set();
    for (let ref of references) {
        set.add(ref.getAttr("w:val"));
    }
    return set;
}
function populateStyles(styles, table) {
    for (let styleId of styles) {
        let style = table.get(styleId);
        if (!style) {
            throw new Error("Style id " + styleId + " not found");
        }
        visitStyleCrossReferences(style, (node) => {
            styles.add(node.getAttr("w:val"));
        });
    }
}
function getUsedStylesDeep(doc, styleTable, requiredStyles = []) {
    let usedStyles = getUsedStyles(doc);
    for (let requiredStyle of requiredStyles) {
        usedStyles.add(requiredStyle);
    }
    do {
        let size = usedStyles.size;
        populateStyles(usedStyles, styleTable);
        if (usedStyles.size == size)
            break;
    } while (true);
    return usedStyles;
}
function getStyleTable(styles) {
    let table = new Map();
    styles.getChild("w:styles").visitChildren("w:style", (style) => {
        table.set(style.getAttr("w:styleId"), style.shallowCopy());
    });
    return table;
}
function getStyleIdsByName(document) {
    return getStyleIdsByNameFromDefs(extractStyleDefs(document));
}
function getStyleIdsByNameFromDefs(styles) {
    let table = new Map();
    for (let style of styles) {
        let nameNode = style.getChild("w:name");
        if (nameNode) {
            table.set(nameNode.getAttr("w:val"), style.getAttr("w:styleId"));
        }
    }
    return table;
}
function addCollisionPatch(mappingTable, styleId) {
    let newId = "template-" + mappingTable.size.toString();
    mappingTable.set(styleId, newId);
    return newId;
}
function getMappingTable(usedStyles) {
    let mappingTable = new Map;
    for (let style of usedStyles) {
        addCollisionPatch(mappingTable, style);
    }
    return mappingTable;
}
function appendStyles(target, defs) {
    let styles = target.getChild("w:styles");
    for (let def of defs) {
        styles.pushChild(def);
    }
}
function applyListStyles(doc, styles) {
    let stack = [];
    let currentState = undefined;
    let newStyles = new Map();
    let lastId = 10000;
    doc.visitSubtree((node) => {
        let tagName = node.getTagName();
        if (tagName === "w:pPr" && currentState) {
            // Remove any old pStyle and add our own
            node.removeChildren("w:pStyle");
            node.pushChild(XML.Node.build("w:pStyle").setAttr("w:val", styles[currentState.listStyle].styleName));
        }
        if (tagName === "w:numId" && currentState) {
            node.setAttr("w:val", String(currentState.numId));
        }
        if (tagName === XML.keys.comment) {
            let commentValue = node.getComment();
            // Switch between ordered list and bullet list
            // if comment is detected
            if (commentValue.indexOf("ListMode OrderedList") != -1) {
                stack.push(currentState);
                currentState = {
                    numId: lastId++,
                    listStyle: "OrderedList"
                };
                newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId);
            }
            if (commentValue.indexOf("ListMode BulletList") != -1) {
                stack.push(currentState);
                currentState = {
                    numId: lastId++,
                    listStyle: "BulletList"
                };
                newStyles.set(String(currentState.numId), styles[currentState.listStyle].numId);
            }
            if (commentValue.indexOf("ListMode None") != -1) {
                currentState = stack[stack.length - 1];
                stack.pop();
            }
        }
        return true;
    });
    return newStyles;
}
function removeCollidedStyles(styles, collisions) {
    let ignored = 0;
    let newChildren = [];
    styles.getChild("w:styles").visitChildren((style) => {
        if (style.getTagName() !== "w:style" || !collisions.has(style.getAttr("w:styleId"))) {
            newChildren.push(style.shallowCopy());
        }
    });
    styles.getChild("w:styles").clearChildren().insertChildren(newChildren);
}
function copyLatentStyles(source, target) {
    let sourceLatentStyles = source.getChild("w:styles").getChild("w:latentStyles");
    let targetLatentStyles = target.getChild("w:styles").getChild("w:latentStyles");
    targetLatentStyles.assign(sourceLatentStyles);
}
function copyDocDefaults(source, target) {
    let sourceDocDefaults = source.getChild("w:styles").getChild("w:docDefaults");
    let targetDocDefaults = target.getChild("w:styles").getChild("w:docDefaults");
    targetDocDefaults.assign(sourceDocDefaults);
}
async function copyFile(source, target, path) {
    target.file(path, await source.file(path).async("arraybuffer"));
}
function addNewNumberings(targetNumberingParsed, newListStyles) {
    let numberingTag = targetNumberingParsed.getChild("w:numbering");
    // <w:num w:numId="newNum">
    //   <w:abstractNumId w:val="oldNum"/>
    // </w:num>
    for (let [newNum, oldNum] of newListStyles) {
        let overrides = [];
        for (let i = 0; i < 9; i++) {
            overrides.push(XML.Node.build("w:lvlOverride")
                .setAttr("w:ilvl", String(i))
                .insertChildren([
                XML.Node.build("w:startOverride").setAttr("w:val", "1")
            ]));
        }
        numberingTag.pushChild(XML.Node.build("w:num")
            .setAttr("w:numId", newNum)
            .insertChildren([
            XML.Node.build("w:abstractNumId").setAttr("w:val", oldNum),
            ...overrides
        ]));
    }
}
function addContentType(contentTypes, partName, contentType) {
    contentTypes.getChild("Types").pushChild(XML.Node.build("Override")
        .setAttr("PartName", partName)
        .setAttr("ContentType", contentType));
}
function transferRels(source, target) {
    let sourceRels = source.getChild("Relationships");
    let targetRels = target.getChild("Relationships");
    let presentIds = new Map();
    let idMap = new Map();
    targetRels.visitChildren((rel) => {
        presentIds.set(rel.getAttr("Target"), rel.getAttr("Id"));
    });
    let newIdCounter = 0;
    sourceRels.visitChildren((rel) => {
        if (presentIds.has(rel.getAttr("Target"))) {
            idMap.set(rel.getAttr("Id"), presentIds.get(rel.getAttr("Target")));
        }
        else {
            let newId = "template-id-" + (newIdCounter++);
            let relCopy = rel.deepCopy();
            relCopy.setAttr("Id", newId);
            targetRels.pushChild(relCopy);
            idMap.set(rel.getAttr("Id"), newId);
        }
    });
    return idMap;
}
function getParagraphText(paragraph) {
    let result = "";
    paragraph.visitSubtree("w:t", (node) => {
        result += getRawText(node);
    });
    return result;
}
function getRawText(tag) {
    let result = "";
    tag.visitSubtree(XML.keys.text, (node) => {
        result += node.getText();
    });
    return result;
}
function replaceInlineTemplate(node, template, value) {
    if (value === "@none") {
        let i = findParagraphWithPattern(node, template, 0);
        for (; i !== null; i = findParagraphWithPattern(node, template, i)) {
            node.removeChild([i]);
            i = i - 1;
        }
    }
    else {
        replaceStringTemplate(node, template, value);
    }
}
function replaceStringTemplate(tag, template, value) {
    tag.visitSubtree(XML.keys.text, (node) => {
        node.setText(node.getText().replace(template, value));
    });
}
function findParagraphWithPattern(node, pattern, startIndex = 0) {
    let found = null;
    node.visitChildren((rel, path) => {
        let text = getParagraphText(rel);
        if (text.indexOf(pattern) !== -1) {
            found = path;
            return false;
        }
        return true;
    }, startIndex);
    return found;
}
function findParagraphWithPatternStrict(body, pattern, startIndex = 0) {
    let paragraphIndex = findParagraphWithPattern(body, pattern, startIndex);
    if (paragraphIndex === null) {
        throw new Error(`The template document should have pattern ${pattern}`);
    }
    let text = getParagraphText(body.getChild([paragraphIndex]));
    if (text != pattern) {
        throw new Error(`The ${pattern} pattern should be the only text of the paragraph`);
    }
    return paragraphIndex;
}
function templateReplaceBodyContents(templateBody, body) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{body}}}");
    let children = body.getChildren();
    templateBody.removeChild([paragraphIndex]);
    templateBody.insertChildren(children, [paragraphIndex]);
}
function clearParagraphContents(paragraph) {
    paragraph.removeChildren("w:r");
}
function templateAuthorList(templateBody, meta) {
    let authors = meta.getSection("ispras_templates.authors").asArray();
    for (let language of exports.languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{authors_${language}}}}`);
        let newParagraphs = [];
        let authorIndex = 1;
        let template = templateBody.getChild([paragraphIndex]);
        for (let author of authors) {
            let newParagraph = template.deepCopy();
            clearParagraphContents(newParagraph);
            let name = author.getString("name_" + language);
            let orcid = author.getString("orcid");
            let email = author.getString("email");
            let indexLine = String(authorIndex);
            let authorLine = `${name}, ORCID: ${orcid}, <${email}>`;
            newParagraph.pushChild(OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()]));
            newParagraph.pushChild(OXML.buildParagraphTextTag(authorLine));
            newParagraphs.push(newParagraph);
            authorIndex++;
        }
        templateBody.removeChild([paragraphIndex]);
        templateBody.insertChildren(newParagraphs, [paragraphIndex]);
    }
    for (let language of exports.languages) {
        let paragraphIndex = findParagraphWithPatternStrict(templateBody, `{{{organizations_${language}}}}`);
        let organizations = meta.getSection("ispras_templates.organizations_" + language).asArray();
        let newParagraphs = [];
        let orgIndex = 1;
        let template = templateBody.getChild([paragraphIndex]);
        for (let organization of organizations) {
            let newParagraph = template.deepCopy();
            clearParagraphContents(newParagraph);
            let indexLine = String(orgIndex);
            newParagraph.pushChild(OXML.buildParagraphTextTag(indexLine, [OXML.buildSuperscriptTextStyle()]));
            newParagraph.pushChild(OXML.buildParagraphTextTag(organization.getString()));
            orgIndex++;
        }
        templateBody.removeChild([paragraphIndex]);
        templateBody.insertChildren(newParagraphs, [paragraphIndex]);
    }
}
function templateReplaceLinks(templateBody, meta, listRules) {
    let litListRule = listRules["LitList"];
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{links}}}");
    let links = meta.getSection("ispras_templates.links").asArray();
    let newParagraphs = [];
    for (let link of links) {
        let newParagraph = OXML.buildParagraphWithStyle(litListRule.styleName);
        let style = newParagraph.getChild("w:pPr");
        style.pushChild(OXML.buildNumPr("0", litListRule.numId));
        newParagraph.pushChild(OXML.buildParagraphTextTag(link.getString()));
        newParagraphs.push(newParagraph);
    }
    templateBody.removeChild([paragraphIndex]);
    templateBody.insertChildren(newParagraphs, [paragraphIndex]);
}
function templateReplaceAuthorsDetail(templateBody, meta) {
    let paragraphIndex = findParagraphWithPatternStrict(templateBody, "{{{authors_detail}}}");
    let authors = meta.getSection("ispras_templates.authors").asArray();
    let newParagraphs = [];
    let template = templateBody.getChild([paragraphIndex]);
    for (let author of authors) {
        for (let language of exports.languages) {
            let newParagraph = template.deepCopy();
            let line = author.getString("details_" + language);
            clearParagraphContents(newParagraph);
            newParagraph.pushChild(OXML.buildParagraphTextTag(line));
            newParagraphs.push(newParagraph);
        }
    }
    templateBody.removeChild([paragraphIndex]);
    templateBody.insertChildren(newParagraphs, [paragraphIndex]);
}
function replacePageHeaders(headers, meta) {
    let header_ru = meta.getString("ispras_templates.page_header_ru");
    let header_en = meta.getString("ispras_templates.page_header_en");
    if (header_ru === "@use_citation") {
        header_ru = meta.getString("ispras_templates.for_citation_ru");
    }
    if (header_en === "@use_citation") {
        header_en = meta.getString("ispras_templates.for_citation_en");
    }
    for (let header of headers) {
        replaceInlineTemplate(header, `{{{page_header_ru}}}`, header_ru);
        replaceInlineTemplate(header, `{{{page_header_en}}}`, header_en);
    }
}
function replaceTemplates(template, body, meta) {
    let templateCopy = template.deepCopy();
    let templateBody = OXML.getDocumentBody(templateCopy);
    templateReplaceBodyContents(templateBody, body);
    templateAuthorList(templateBody, meta);
    let templates = ["header", "abstract", "keywords", "for_citation", "acknowledgements"];
    for (let template of templates) {
        for (let language of exports.languages) {
            let template_lang = template + "_" + language;
            let value = meta.getString("ispras_templates." + template_lang);
            replaceInlineTemplate(templateBody, `{{{${template_lang}}}}`, value);
        }
    }
    templateReplaceAuthorsDetail(templateBody, meta);
    return templateCopy;
}
function setXmlns(xml, xmlns) {
    const document = xml.getChild("w:document");
    for (let [key, value] of xmlns) {
        document.setAttr(key, value);
    }
}
function patchRelIds(doc, map) {
    doc.visitSubtree((node) => {
        for (let attr of ["r:id", "r:embed"]) {
            let relId = node.getAttr(attr);
            if (relId && map.has(relId)) {
                node.setAttr(attr, map.get(relId));
            }
        }
        return true;
    });
}
async function fixDocxStyles(sourcePath, targetPath, meta) {
    let resourcesDir = path.dirname(process.argv[1]) + "/../resources";
    // Load the source and target documents
    let target = await JSZip.loadAsync(fs.readFileSync(sourcePath));
    let source = await JSZip.loadAsync(fs.readFileSync(resourcesDir + '/isp-reference.docx'));
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
    let targetContentTypesParsed = XML.Node.fromXmlString(targetContentTypesXML);
    let targetDocumentRelsParsed = XML.Node.fromXmlString(targetDocumentRelsXML);
    let sourceDocumentRelsParsed = XML.Node.fromXmlString(sourceDocumentRelsXML);
    let sourceStylesParsed = XML.Node.fromXmlString(sourceStylesXML);
    let targetStylesParsed = XML.Node.fromXmlString(targetStylesXML);
    let sourceDocParsed = XML.Node.fromXmlString(sourceDocXML);
    let targetDocParsed = XML.Node.fromXmlString(targetDocXML);
    let targetNumberingParsed = XML.Node.fromXmlString(targetNumberingXML);
    let sourceHeader1Parsed = XML.Node.fromXmlString(sourceHeader1);
    let sourceHeader2Parsed = XML.Node.fromXmlString(sourceHeader2);
    let sourceHeader3Parsed = XML.Node.fromXmlString(sourceHeader3);
    copyLatentStyles(sourceStylesParsed, targetStylesParsed);
    copyDocDefaults(sourceStylesParsed, targetStylesParsed);
    let targetStylesNamesToId = getStyleIdsByName(targetStylesParsed);
    let sourceStylesNamesToId = getStyleIdsByName(sourceStylesParsed);
    let sourceStyleTable = getStyleTable(sourceStylesParsed);
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
    ].map(name => sourceStylesNamesToId.get(name)));
    let mappingTable = getMappingTable(usedStyles);
    patchStyleDefinitions(sourceDocParsed, sourceStylesParsed, mappingTable);
    patchStyleUseReferences(sourceDocParsed, sourceStylesParsed, mappingTable);
    let extractedDefs = extractTemplateStyleDefs(sourceStylesParsed);
    let extractedStyleIdsByName = getStyleIdsByNameFromDefs(extractedDefs);
    let stylePatch = new Map([
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
    ]);
    let stylesToRemove = new Set([
        "Heading4",
        "Heading5",
        "Heading6",
        "Heading7",
        "Heading8",
        "Heading9",
    ]);
    for (let possibleCollision of extractedStyleIdsByName) {
        let sourceStyleName = possibleCollision[0];
        let sourceStyleId = possibleCollision[1];
        if (targetStylesNamesToId.has(sourceStyleName)) {
            let targetStyleId = targetStylesNamesToId.get(sourceStyleName);
            if (!stylePatch.has(targetStyleId)) {
                stylePatch.set(targetStyleId, sourceStyleId);
            }
            stylesToRemove.add(targetStyleId);
        }
    }
    removeCollidedStyles(targetStylesParsed, stylesToRemove);
    appendStyles(targetStylesParsed, extractedDefs);
    patchStyleUseReferences(targetDocParsed, targetStylesParsed, stylePatch);
    let patchRules = {
        "OrderedList": { styleName: extractedStyleIdsByName.get("ispNumList"), numId: "33" },
        "BulletList": { styleName: extractedStyleIdsByName.get("ispList1"), numId: "43" },
        "LitList": { styleName: extractedStyleIdsByName.get("ispLitList"), numId: "80" },
    };
    let newListStyles = applyListStyles(targetDocParsed, patchRules);
    setXmlns(sourceDocParsed, properDocXmlns);
    let relMap = transferRels(sourceDocumentRelsParsed, targetDocumentRelsParsed);
    patchRelIds(sourceDocParsed, relMap);
    targetDocParsed = replaceTemplates(sourceDocParsed, OXML.getDocumentBody(targetDocParsed), meta);
    templateReplaceLinks(OXML.getDocumentBody(targetDocParsed), meta, patchRules);
    addNewNumberings(targetNumberingParsed, newListStyles);
    replacePageHeaders([sourceHeader1Parsed, sourceHeader2Parsed, sourceHeader3Parsed], meta);
    addContentType(targetContentTypesParsed, "/word/footer1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
    addContentType(targetContentTypesParsed, "/word/footer2.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
    addContentType(targetContentTypesParsed, "/word/footer3.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml");
    addContentType(targetContentTypesParsed, "/word/header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
    addContentType(targetContentTypesParsed, "/word/header2.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
    addContentType(targetContentTypesParsed, "/word/header3.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml");
    await copyFile(source, target, "word/_rels/header1.xml.rels");
    await copyFile(source, target, "word/_rels/header2.xml.rels");
    await copyFile(source, target, "word/_rels/header3.xml.rels");
    await copyFile(source, target, "word/_rels/footer1.xml.rels");
    await copyFile(source, target, "word/_rels/footer2.xml.rels");
    await copyFile(source, target, "word/_rels/footer3.xml.rels");
    await copyFile(source, target, "word/footer1.xml");
    await copyFile(source, target, "word/footer2.xml");
    await copyFile(source, target, "word/footer3.xml");
    await copyFile(source, target, "word/footnotes.xml");
    await copyFile(source, target, "word/theme/theme1.xml");
    await copyFile(source, target, "word/fontTable.xml");
    await copyFile(source, target, "word/settings.xml");
    await copyFile(source, target, "word/webSettings.xml");
    await copyFile(source, target, "word/media/image1.png");
    target.file("word/header1.xml", sourceHeader1Parsed.toXmlString());
    target.file("word/header2.xml", sourceHeader2Parsed.toXmlString());
    target.file("word/header3.xml", sourceHeader3Parsed.toXmlString());
    target.file("word/_rels/document.xml.rels", targetDocumentRelsParsed.toXmlString());
    target.file("[Content_Types].xml", targetContentTypesParsed.toXmlString());
    target.file("word/numbering.xml", targetNumberingParsed.toXmlString());
    target.file("word/styles.xml", targetStylesParsed.toXmlString());
    target.file("word/document.xml", targetDocParsed.toXmlString());
    fs.writeFileSync(targetPath, await target.generateAsync({ type: "uint8array" }));
}
function fixCompactLists(list) {
    // For compact list, 'para' is replaced with 'plain'.
    // Compact lists were not mentioned in the
    // guidelines, so get rid of them
    for (let i = 0; i < list.c.length; i++) {
        let element = list.c[i];
        if (typeof element[0] === "object" && element[0].t === "Plain") {
            element[0].t = "Para";
        }
        list.c[i] = getPatchedMetaElement(list.c[i]);
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
    ];
}
function getImageCaption(content) {
    let paragraph = XML.Node.build("w:p").insertChildren([
        XML.Node.build("w:pPr").insertChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", "ImageCaption"),
            XML.Node.build("w:contextualSpacing").setAttr("w:val", "true"),
        ]),
        OXML.buildParagraphTextTag(pandoc.getMetaString(content))
    ]);
    return {
        t: "RawBlock",
        c: ["openxml", paragraph.toXmlString()]
    };
}
function getListingCaption(content) {
    let elements = XML.Node.build("w:p").insertChildren([
        XML.Node.build("w:pPr").insertChildren([
            XML.Node.build("w:pStyle").setAttr("w:val", "BodyText"),
            XML.Node.build("w:jc").setAttr("w:val", "left"),
        ]),
        OXML.buildParagraphTextTag(pandoc.getMetaString(content), [
            XML.Node.build("w:i"),
            XML.Node.build("w:iCs"),
            XML.Node.build("w:sz").setAttr("w:val", "18"),
            XML.Node.build("w:szCs").setAttr("w:val", "18"),
        ])
    ]);
    return {
        t: "RawBlock",
        c: ["openxml", elements.toXmlString()]
    };
}
function getPatchedMetaElement(element) {
    if (Array.isArray(element)) {
        let newArray = [];
        for (let i = 0; i < element.length; i++) {
            let patched = getPatchedMetaElement(element[i]);
            if (Array.isArray(patched) && !Array.isArray(element[i])) {
                newArray.push(...patched);
            }
            else {
                newArray.push(patched);
            }
        }
        return newArray;
    }
    if (typeof element !== "object" || !element) {
        return element;
    }
    let type = element.t;
    let value = element.c;
    if (type === 'Div') {
        let content = value[1];
        let classes = value[0][1];
        if (classes) {
            if (classes.includes("img-caption")) {
                return getImageCaption(content);
            }
            if (classes.includes("table-caption") || classes.includes("listing-caption")) {
                return getListingCaption(content);
            }
        }
    }
    else if (type === 'BulletList' || type === 'OrderedList') {
        return fixCompactLists(element);
    }
    for (let key of Object.getOwnPropertyNames(element)) {
        // Be safe from prototype pollution
        if (key === "__proto__")
            continue;
        element[key] = getPatchedMetaElement(element[key]);
    }
    return element;
}
async function generatePandocDocx(source, target) {
    let markdown = await fs.promises.readFile(source, "utf-8");
    let meta = await pandoc.pandoc(markdown, ["-f", "markdown", "-t", "json", ...pandocFlags]);
    let metaParsed = JSON.parse(meta);
    metaParsed.blocks = getPatchedMetaElement(metaParsed.blocks);
    await pandoc.pandoc(JSON.stringify(metaParsed), ["-f", "json", "-t", "docx", "-o", target]);
    return pandoc_1.DocumentMeta.fromPandocMeta(metaParsed.meta);
}
async function main() {
    let argv = process.argv;
    if (argv.length < 4) {
        console.log("Usage: main.js <source> <target>");
        process.exit(1);
    }
    let source = argv[2];
    let target = argv[3];
    let meta = await generatePandocDocx(source, target + ".tmp");
    await fixDocxStyles(target + ".tmp", target, meta).then();
}
main().then();
//# sourceMappingURL=main.js.map