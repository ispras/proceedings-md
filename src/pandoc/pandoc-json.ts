/*
    This code was taken from https://github.com/mvhenderson/pandoc-filter-node
    License for pandoc-filter-node:

    Copyright (c) 2014 Mike Henderson

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.
 */

export type PandocJson = {
    blocks: Block[];
    "pandoc-api-version": number[];
    meta: PandocMetaMap;
};

/** list of key-value attributes */
export type AttrList = Array<[string, string]>;

/** [id, classes, list of key-value attributes] */
export type Attr = [string, Array<string>, AttrList];

export type MathType = { t: "DisplayMath" | "InlineMath" };
export type QuoteType = { t: "SingleQuote" | "DoubleQuote" };

/** [url, title] */
export type Target = [string, string];
/** output file format */
export type Format = string;

export type CitationMode = {
    t: "AuthorInText" | "SuppressAuthor" | "NormalCitation";
};

export type Citation = {
    citationId: string;
    citationPrefix: Array<Inline>;
    citationSuffix: Array<Inline>;
    citationMode: CitationMode;
    citationNoteNum: number;
    citationHash: number;
};

export type ListNumberStyle = {
    t:
        | "DefaultStyle"
        | "Example"
        | "Decimal"
        | "LowerRoman"
        | "UpperRoman"
        | "LowerAlpha"
        | "UpperAlpha";
};

export type ListNumberDelim = {
    t: "DefaultDelim" | "Period" | "OneParen" | "TwoParens";
};

export type ListAttributes = [number, ListNumberStyle, ListNumberDelim];

export type Alignment = {
    t: "AlignLeft" | "AlignRight" | "AlignCenter" | "AlignDefault";
};

export type ColWidth = {
    t: "ColWidth";
    c: number;
};

export type ColWidthDefault = {
    t: "ColWidthDefault";
};

export type Caption = [
    Array<Inline>, // short
    Array<Block>, // long
];

export type TableCaption = Caption;

export type FigureCaption = Caption;

export type TableHead = [Attr, Array<TableRow>];

export type TableBody = [
    [
        Attr,
        number, // row head columns
        Array<TableRow>, // intermediate head
        Array<TableRow>, // body rows
    ],
];

export type TableFoot = [Attr, Array<Block>];

export type TableRow = [Attr, Array<TableCell>];

export type TableCell = [
    Attr,
    Alignment,
    number, // row span
    number, // col span
    Array<Block>,
];

export type MetaElementMap = {
    // Inline
    Str: string;
    Emph: Array<Inline>;
    Strong: Array<Inline>;
    Strikeout: Array<Inline>;
    Superscript: Array<Inline>;
    Subscript: Array<Inline>;
    SmallCaps: Array<Inline>;
    Quoted: [QuoteType, Array<Inline>];
    Cite: [Array<Citation>, Array<Inline>];
    Code: [Attr, string];
    Space: undefined;
    SoftBreak: undefined;
    LineBreak: undefined;
    Math: [MathType, string];
    RawInline: [Format, string];
    Link: [Attr, Array<Inline>, Target];
    Image: [Attr, Array<Inline>, Target];
    Note: Array<Block>;
    Span: [Attr, Array<Inline>];

    // Block
    Plain: Array<Inline>;
    Para: Array<Inline>;
    LineBlock: Array<Array<Inline>>;
    CodeBlock: [Attr, string];
    RawBlock: [Format, string];
    BlockQuote: Array<Block>;
    OrderedList: [ListAttributes, Array<Array<Block>>];
    BulletList: Array<Array<Block>>;
    DefinitionList: Array<[Array<Inline>, Array<Array<Block>>]>;
    Header: [number, Attr, Array<Inline>];
    HorizontalRule: undefined;
    Table: [
        Attr,
        TableCaption,
        Array<[Alignment, ColWidth | ColWidthDefault]>,
        TableHead,
        TableBody,
        TableFoot,
    ];
    Figure: [Attr, FigureCaption, Array<Block>];
    Div: [Attr, Array<Block>];
    Null: undefined;
};
export type MetaElementType = keyof MetaElementMap;

export type MetaElement<A extends MetaElementType> = { t: A; c: MetaElementMap[A] };

export type AnyElement = Inline | Block;

export type Inline =
    | MetaElement<"Str">
    | MetaElement<"Emph">
    | MetaElement<"Strong">
    | MetaElement<"Strikeout">
    | MetaElement<"Superscript">
    | MetaElement<"Subscript">
    | MetaElement<"SmallCaps">
    | MetaElement<"Quoted">
    | MetaElement<"Cite">
    | MetaElement<"Code">
    | MetaElement<"Space">
    | MetaElement<"SoftBreak">
    | MetaElement<"LineBreak">
    | MetaElement<"Math">
    | MetaElement<"RawInline">
    | MetaElement<"Link">
    | MetaElement<"Image">
    | MetaElement<"Note">
    | MetaElement<"Span">;

export type Block =
    | MetaElement<"Plain">
    | MetaElement<"Para">
    | MetaElement<"LineBlock">
    | MetaElement<"CodeBlock">
    | MetaElement<"RawBlock">
    | MetaElement<"BlockQuote">
    | MetaElement<"OrderedList">
    | MetaElement<"BulletList">
    | MetaElement<"DefinitionList">
    | MetaElement<"Header">
    | MetaElement<"HorizontalRule">
    | MetaElement<"Table">
    | MetaElement<"Figure">
    | MetaElement<"Div">
    | MetaElement<"Null">;

/** meta information about document, mostly from markdown frontmatter
 * https://hackage.haskell.org/package/pandoc-types-1.20/docs/Text-Pandoc-Definition.html#t:MetaValue
 */
export type PandocMetaValue =
    | { t: "MetaMap"; c: PandocMetaMap }
    | { t: "MetaList"; c: Array<PandocMetaValue> }
    | { t: "MetaBool"; c: boolean }
    | { t: "MetaInlines"; c: Inline[] }
    | { t: "MetaString"; c: string }
    | { t: "MetaBlocks"; c: Block[] };
export type PandocMetaMap = Record<string, PandocMetaValue>;

function isElement(x: unknown): x is AnyElement {
    return (typeof x === "object" && x && "t" in x) || false;
}

export function walkPandocElement(
    object: any,
    action: (ele: AnyElement) => void | AnyElement | Array<AnyElement>
) {
    if (Array.isArray(object)) {
        let array: any[] = [];

        for (const element of object) {
            if (!isElement(element)) {
                array.push(walkPandocElement(element, action));
                continue
            }

            let replacement = action(element)

            if(replacement) {
                if (Array.isArray(replacement)) {
                    array.push(...replacement);
                } else {
                    array.push(replacement);
                }
            } else {
                array.push(walkPandocElement(element, action));
            }
        }
        return array;
    }

    if (typeof object === "object" && object !== null) {
        let result: any = {};

        for (const key of Object.getOwnPropertyNames(object)) {
            if(key === "__proto__") continue
            result[key] = walkPandocElement((object as any)[key], action);
        }

        return result;
    }

    return object;
}

export function metaElementToSource(value: any) {
    let result: string[] = []

    walkPandocElement(value, (child) => {
        if (child.t === "Str") result.push(child.c as string)
        else if (child.t === "Strong") result.push("__" + metaElementToSource(child.c) + "__")
        else if (child.t === "Emph") result.push("_" + metaElementToSource(child.c) + "_")
        else if (child.t === "Space") result.push(" ")
        else if (child.t === "LineBreak") result.push("\n")
        else if (child.t === "Code") result.push("`" + child.c[1] + "`");
        else return undefined
        return child
    })

    return result.join("")
}