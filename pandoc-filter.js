#!/usr/bin/env node
"use strict";

let fs = require("fs")
let pandoc = require("pandoc-filter");
const {XMLBuilder} = require("fast-xml-parser");

let builder = new XMLBuilder({
    ignoreAttributes: false,
    alwaysCreateTextNode: true,
    attributeNamePrefix: "",
    preserveOrder: true
})

function extractText(content, met = new Set()) {
    if (Array.isArray(content)) {
        let result = ""
        for (let child of content) {
            result += extractText(child)
        }
        return result
    } else if (content.t === "Str") {
        return content.c
    } else if (content.t === "Space") {
        return " "
    } else if (content.c) {
        return extractText(content.c)
    }
    return ""
}

function fixCompactLists(list) {
    // For compact list, 'para' is replaced with 'plain'.
    // Compact lists were not mentioned in the
    // guidelines, so get rid of them

    for(let element of list.c) {
        if(typeof element[0] === "object" && element[0].t === "Plain") {
            element[0].t = "Para"
        }
    }

    return [
        pandoc.RawBlock("openxml", `<!-- ListMode ${list.t} -->`),
        list,
        pandoc.RawBlock("openxml", `<!-- ListMode None -->`),]
}

function getImageCaption(content) {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ":@": {"w:val": "ImageCaption"}
                }, {
                    "w:contextualSpacing": [],
                    ":@": {
                        "w:val": "true"
                    }
                }]
        },
        {
            "w:r": [{
                "w:t": [{
                    "#text": extractText(content)
                }],
                ":@": {
                    "xml:space": "preserve"
                }
            }]
        }
    ];
    return new pandoc.RawBlock("openxml", `<w:p>${builder.build(elements)}</w:p>`);
}

function getListingCaption(content) {
    let elements = [
        {
            "w:pPr": [
                {
                    "w:pStyle": [],
                    ":@": {"w:val": "BodyText"}
                }, {
                    "w:jc": [],
                    ":@": {"w:val": "left"}
                },
            ]
        },
        {
            "w:r": [
                {
                    "w:rPr": [
                        {"w:i": []},
                        {"w:iCs": []},
                        {"w:sz": [], ":@": {"w:val": "18"}},
                        {"w:szCs": [], ":@": {"w:val": "18"}},
                    ]}, {
                    "w:t": [{
                        "#text": extractText(content)
                    }
                ],
                ":@": {
                    "xml:space": "preserve"
                }
            }]
        }
    ];
    return new pandoc.RawBlock("openxml", `<w:p>${builder.build(elements)}</w:p>`);
}

async function action(element, format, meta) {
    let type = element.t
    let value = element.c

    if (format !== "docx")
        return;

    if (type === 'Div') {
        let content = value[1];
        let classes = value[0][1];
        if (!classes)
            return;

        if (classes.includes("img-caption")) {
            return getImageCaption(content)
        }

        if (classes.includes("table-caption") || classes.includes("listing-caption")) {
            return getListingCaption(content)
        }
    } else if (type === 'BulletList' || type === 'OrderedList') {
        return fixCompactLists(element)
    }

}

pandoc.stdio(action);