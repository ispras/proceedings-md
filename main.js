const fs = require('fs');
const docx = require('docx');
const JSZip = require("jszip");
const { XMLParser, XMLBuilder, XMLValidator } = require("fast-xml-parser");
function getChildTag(styles, name) {
    for (let child of styles) {
        if (child[name]) {
            return child;
        }
    }
}
function getStyleCrossReferences(styles) {
    let result = [];
    for (let style of getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"])
            continue;
        let basedOnTag = getChildTag(style["w:style"], "w:basedOn");
        if (basedOnTag)
            result.push(basedOnTag[":@"]);
        let linkTag = getChildTag(style["w:style"], "w:link");
        if (linkTag)
            result.push(linkTag[":@"]);
        let nextTag = getChildTag(style["w:style"], "w:next");
        if (nextTag)
            result.push(nextTag[":@"]);
    }
    return result;
}
function getDocStyleUseReferences(doc, result = [], met = new Set()) {
    if (!doc || typeof doc !== "object" || met.has(doc)) {
        return result;
    }
    met.add(doc);
    for (let key of Object.getOwnPropertyNames(doc)) {
        if (key === "w:pStyle" || key == "w:rStyle") {
            result.push(doc[":@"]);
        }
        else {
            result = getDocStyleUseReferences(doc[key], result, met);
        }
    }
    return result;
}
function extractStyleDefs(styles, map) {
    let result = [];
    for (let style of getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"])
            continue;
        if (map.has(style[":@"]["w:styleId"])) {
            let copy = JSON.parse(JSON.stringify(style));
            result.push(copy);
        }
    }
    return result;
}
function patchStyleDefs(styles, map) {
    let result = [];
    for (let style of styles) {
        if (map.has(style[":@"]["w:styleId"])) {
            style[":@"]["w:styleId"] = map.get(style[":@"]["w:styleId"]);
            let basedOnTag = getChildTag(style["w:style"], "w:basedOn");
            if (basedOnTag)
                basedOnTag[":@"]["w:val"] = map.get(basedOnTag[":@"]["w:val"]);
            let linkTag = getChildTag(style["w:style"], "w:link");
            if (linkTag)
                linkTag[":@"]["w:val"] = map.get(linkTag[":@"]["w:val"]);
            let nextTag = getChildTag(style["w:style"], "w:next");
            if (nextTag)
                nextTag[":@"]["w:val"] = map.get(nextTag[":@"]["w:val"]);
            result.push(style);
        }
    }
    return result;
}
function patchStyleUseReferences(doc, styles, map) {
    let docReferences = getDocStyleUseReferences(doc);
    let crossReferences = getStyleCrossReferences(styles);
    for (let ref of docReferences.concat(crossReferences)) {
        if (!map.has(ref["w:val"]))
            continue;
        ref["w:val"] = map.get(ref["w:val"]);
    }
}
function getUsedStyles(doc) {
    let references = getDocStyleUseReferences(doc);
    let set = new Set();
    for (let ref of references) {
        set.add(ref["w:val"]);
    }
    return set;
}
function populateStyles(styles, table) {
    for (let styleId of styles) {
        let style = table.get(styleId);
        if (!style) {
            throw new Error("Style id " + styleId + " not found");
        }
        let basedOnTag = getChildTag(style["w:style"], "w:basedOn");
        if (basedOnTag)
            styles.add(basedOnTag[":@"]["w:val"]);
        let linkTag = getChildTag(style["w:style"], "w:link");
        if (linkTag)
            styles.add(linkTag[":@"]["w:val"]);
        let nextTag = getChildTag(style["w:style"], "w:next");
        if (nextTag)
            styles.add(nextTag[":@"]["w:val"]);
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
    for (let style of getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"])
            continue;
        table.set(style[":@"]["w:styleId"], style);
    }
    return table;
}
function getStyleIdsByNameFromDefs(styles) {
    let table = new Map();
    for (let style of styles) {
        if (!style["w:style"])
            continue;
        let nameNode = getChildTag(style["w:style"], "w:name");
        if (nameNode) {
            table.set(nameNode[":@"]["w:val"], style[":@"]["w:styleId"]);
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
    let styles = getChildTag(target, "w:styles")["w:styles"];
    for (let def of defs) {
        styles.push(def);
    }
}
function applyListStyles(doc, styles) {
    let stack = [];
    let currentState = undefined;
    let met = new Set();
    let newStyles = new Map();
    let lastId = 10000;
    const walk = (doc) => {
        if (!doc || typeof doc !== "object" || met.has(doc)) {
            return;
        }
        met.add(doc);
        for (let key of Object.getOwnPropertyNames(doc)) {
            walk(doc[key]);
            if (key === "w:pPr" && currentState) {
                // Remove any old pStyle and add our own
                for (let i = 0; i < doc[key].length; i++) {
                    if (doc[key][i]["w:pStyle"]) {
                        doc[key].splice(i, 1);
                        i--;
                    }
                }
                doc[key].unshift({
                    "w:pStyle": {},
                    ":@": { "w:val": styles[currentState.listStyle].styleName }
                });
            }
            if (key === "w:numId" && currentState) {
                doc[":@"]["w:val"] = String(currentState.numId);
            }
            if (key === "#comment") {
                let commentValue = doc[key][0]["#text"];
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
        }
    };
    walk(doc);
    return newStyles;
}
function removeCollidedStyles(styles, collisions) {
    let ignored = 0;
    let newContents = [];
    for (let style of getChildTag(styles, "w:styles")["w:styles"]) {
        if (!style["w:style"] || !collisions.has(style[":@"]["w:styleId"])) {
            newContents.push(style);
        }
    }
    getChildTag(styles, "w:styles")["w:styles"] = newContents;
}
function copyLatentStyles(source, target) {
    let sourceStyles = getChildTag(source, "w:styles")["w:styles"];
    let targetStyles = getChildTag(target, "w:styles")["w:styles"];
    let sourceLatentStyles = getChildTag(sourceStyles, "w:latentStyles");
    let targetLatentStyles = getChildTag(targetStyles, "w:latentStyles");
    targetLatentStyles["w:latentStyles"] = JSON.parse(JSON.stringify(sourceLatentStyles["w:latentStyles"]));
    if (targetLatentStyles[":@"]) {
        targetLatentStyles[":@"] = JSON.parse(JSON.stringify(sourceLatentStyles[":@"]));
    }
}
function copyDocDefaults(source, target) {
    let sourceStyles = getChildTag(source, "w:styles")["w:styles"];
    let targetStyles = getChildTag(target, "w:styles")["w:styles"];
    let sourceDocDefaults = getChildTag(sourceStyles, "w:docDefaults");
    let targetDocDefaults = getChildTag(targetStyles, "w:docDefaults");
    targetDocDefaults["w:docDefaults"] = JSON.parse(JSON.stringify(sourceDocDefaults["w:docDefaults"]));
    if (sourceDocDefaults[":@"]) {
        targetDocDefaults[":@"] = JSON.parse(JSON.stringify(sourceDocDefaults[":@"]));
    }
}
function copySectPr(source, target) {
    let sourceDocument = getChildTag(source, "w:document")["w:document"];
    let sourceBody = getChildTag(sourceDocument, "w:body")["w:body"];
    let sourceSectPr = getChildTag(sourceBody, "w:sectPr");
    let targetDocument = getChildTag(target, "w:document")["w:document"];
    let targetBody = getChildTag(targetDocument, "w:body")["w:body"];
    let targetSectPr = getChildTag(targetBody, "w:sectPr");
    targetSectPr["w:sectPr"] = JSON.parse(JSON.stringify(sourceSectPr["w:sectPr"]));
    if (sourceSectPr[":@"]) {
        targetSectPr[":@"] = JSON.parse(JSON.stringify(sourceSectPr[":@"]));
    }
}
async function copyFile(source, target, path) {
    target.file(path, await source.file(path).async("arraybuffer"));
}
function addNewNumberings(targetNumberingParsed, newListStyles) {
    let numberingTag = getChildTag(targetNumberingParsed, "w:numbering")["w:numbering"];
    // <w:num w:numId="newNum">
    //   <w:abstractNumId w:val="oldNum"/>
    // </w:num>
    for (let [newNum, oldNum] of newListStyles) {
        let overrides = [];
        for (let i = 0; i < 9; i++) {
            overrides.push({
                "w:lvlOverride": [{
                        "w:startOverride": [],
                        ":@": {
                            "w:val": "1"
                        }
                    }],
                ":@": {
                    "w:ilvl": String(i)
                }
            });
        }
        numberingTag.push({
            "w:num": [{
                    "w:abstractNumId": [],
                    ":@": {
                        "w:val": oldNum
                    }
                }, ...overrides],
            ":@": {
                "w:numId": newNum
            }
        });
    }
}
function addContentType(contentTypes, partName, contentType) {
    let typesTag = getChildTag(contentTypes, "Types")["Types"];
    typesTag.push({
        "Override": [],
        ":@": {
            "PartName": partName,
            "ContentType": contentType
        }
    });
}
function transferRels(source, target) {
    let sourceRels = getChildTag(source, "Relationships")["Relationships"];
    let targetRels = getChildTag(target, "Relationships")["Relationships"];
    let presentIds = new Set();
    for (let rel of targetRels) {
        presentIds.add(rel[":@"]["Id"]);
    }
    for (let rel of sourceRels) {
        if (!presentIds.has(rel[":@"]["Id"])) {
            targetRels.push(JSON.parse(JSON.stringify(rel)));
        }
    }
}
async function copyStyles() {
    // Load the source and target documents
    // const sourceDoc = new Document(fs.readFileSync('isp-reference.docx'));
    let target = await JSZip.loadAsync(fs.readFileSync('main.docx'));
    let source = await JSZip.loadAsync(fs.readFileSync('isp-reference.docx'));
    let sourceStylesXML = await source.file("word/styles.xml").async("string");
    let targetStylesXML = await target.file("word/styles.xml").async("string");
    let sourceDocXML = await source.file("word/document.xml").async("string");
    let targetDocXML = await target.file("word/document.xml").async("string");
    // fs.writeFileSync("source_styles.xml", sourceStyles)
    // fs.writeFileSync("target_styles.xml", sourceStyles)
    // fs.writeFileSync("source_document.xml", sourceStyles)
    // fs.writeFileSync("target_document.xml", sourceStyles)
    // Parse the source styles
    let parser = new XMLParser({
        ignoreAttributes: false,
        alwaysCreateTextNode: true,
        attributeNamePrefix: "",
        preserveOrder: true,
        trimValues: false,
        commentPropName: "#comment"
    });
    let sourceStylesParsed = parser.parse(sourceStylesXML);
    let targetStylesParsed = parser.parse(targetStylesXML);
    let sourceDocParsed = parser.parse(sourceDocXML);
    let targetDocParsed = parser.parse(targetDocXML);
    copyLatentStyles(sourceStylesParsed, targetStylesParsed);
    copyDocDefaults(sourceStylesParsed, targetStylesParsed);
    copySectPr(sourceDocParsed, targetDocParsed);
    let targetStylesNamesToId = getStyleIdsByNameFromDefs(getChildTag(targetStylesParsed, "w:styles")["w:styles"]);
    let sourceStylesNamesToId = getStyleIdsByNameFromDefs(getChildTag(sourceStylesParsed, "w:styles")["w:styles"]);
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
        "ispListing Знак"
    ].map(name => sourceStylesNamesToId.get(name)));
    let mappingTable = getMappingTable(usedStyles);
    let extractedDefs = extractStyleDefs(sourceStylesParsed, mappingTable);
    patchStyleDefs(extractedDefs, mappingTable);
    let extractedStyleIdsByName = getStyleIdsByNameFromDefs(extractedDefs);
    let stylePatch = new Map([
        ["Heading1", extractedStyleIdsByName.get("ispSubHeader-1 level")],
        ["Heading2", extractedStyleIdsByName.get("ispSubHeader-2 level")],
        ["Heading3", extractedStyleIdsByName.get("ispSubHeader-3 level")],
        ["Author", extractedStyleIdsByName.get("ispAuthor")],
        ["AbstractTitle", extractedStyleIdsByName.get("ispAnotation")],
        ["Abstract", extractedStyleIdsByName.get("ispAnotation")],
        ["BlockText", extractedStyleIdsByName.get("ispText_main")],
        ["DefaultParagraphFont", extractedStyleIdsByName.get("DefaultParagraphFont")],
        ["BodyText", extractedStyleIdsByName.get("ispText_main")],
        ["FirstParagraph", extractedStyleIdsByName.get("ispText_main")],
        ["Normal", extractedStyleIdsByName.get("Normal")],
        ["SourceCode", extractedStyleIdsByName.get("ispListing")],
        ["VerbatimChar", extractedStyleIdsByName.get("ispListing Знак")],
        ["ImageCaption", extractedStyleIdsByName.get("ispPicture_sign")],
        // ["Compact",                styleIdsByName.get("ispPicture_sign")],
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
        "BulletList": { styleName: extractedStyleIdsByName.get("ispList1"), numId: "43" }
    };
    let newListStyles = applyListStyles(targetDocParsed, patchRules);
    let targetNumberingXML = await source.file("word/numbering.xml").async("string");
    let targetNumberingParsed = parser.parse(targetNumberingXML);
    addNewNumberings(targetNumberingParsed, newListStyles);
    let builder = new XMLBuilder({
        ignoreAttributes: false,
        alwaysCreateTextNode: true,
        attributeNamePrefix: "",
        preserveOrder: true,
        commentPropName: "#comment"
    });
    let targetContentTypesXML = await target.file("[Content_Types].xml").async("string");
    let targetContentTypesParsed = parser.parse(targetContentTypesXML);
    let targetDocumentRelsXML = await target.file("word/_rels/document.xml.rels").async("string");
    let targetDocumentRelsParsed = parser.parse(targetDocumentRelsXML);
    let sourceDocumentRelsXML = await source.file("word/_rels/document.xml.rels").async("string");
    let sourceDocumentRelsParsed = parser.parse(sourceDocumentRelsXML);
    transferRels(sourceDocumentRelsParsed, targetDocumentRelsParsed);
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
    await copyFile(source, target, "word/header1.xml");
    await copyFile(source, target, "word/header2.xml");
    await copyFile(source, target, "word/header3.xml");
    await copyFile(source, target, "word/footer1.xml");
    await copyFile(source, target, "word/footer2.xml");
    await copyFile(source, target, "word/footer3.xml");
    await copyFile(source, target, "word/footnotes.xml");
    await copyFile(source, target, "word/theme/theme1.xml");
    await copyFile(source, target, "word/fontTable.xml");
    await copyFile(source, target, "word/settings.xml");
    await copyFile(source, target, "word/webSettings.xml");
    await copyFile(source, target, "word/media/image1.png");
    await copyFile(source, target, "word/media/image2.png");
    target.file("word/_rels/document.xml.rels", builder.build(targetDocumentRelsParsed));
    target.file("[Content_Types].xml", builder.build(targetContentTypesParsed));
    target.file("word/numbering.xml", builder.build(targetNumberingParsed));
    target.file("word/styles.xml", builder.build(targetStylesParsed));
    target.file("word/document.xml", builder.build(targetDocParsed));
    fs.writeFileSync("main_styled.docx", await target.generateAsync({ type: "uint8array" }));
}
copyStyles().then();
//# sourceMappingURL=main.js.map