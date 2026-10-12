import { fromMarkdown } from "mdast-util-from-markdown";
import { frontmatterFromMarkdown } from "mdast-util-frontmatter";
import { frontmatter } from "micromark-extension-frontmatter";
import { math } from "micromark-extension-math";
import { mathFromMarkdown } from "mdast-util-math";
import { gfmTableFromMarkdown } from "mdast-util-gfm-table";
import { gfmTable } from "micromark-extension-gfm-table";
import { visit } from "unist-util-visit";
import mdast, { Text } from "mdast";
import DocumentReferences from "./document-references";
import { DocumentJsonMeta } from "../document-json-meta";

declare module "mdast" {
  interface HeadingData {
    id?: string;
  }

  interface ImageData {
    attrs?: { [key: string]: string | number };
  }
}

function addHeaderReferences(tree: mdast.Root) {
  visit(tree, "heading", (node) => {
    let lastChild = node.children[node.children.length - 1];

    if (!lastChild || lastChild.type !== "text") return;

    lastChild = lastChild as Text;

    const match = lastChild.value.match(/\s*{#([^}]+)}$/);
    if (match) {
      node.data = node.data || {};
      node.data.id = match[1];
      lastChild.value = lastChild.value.replace(/\s*{#[^}]+}$/, "").trim();
    }
  });
}

function removePositions(tree: mdast.Root) {
  visit(tree, (node) => {
    node.position = undefined;
  });
}
function addReferences(tree: mdast.Root): DocumentReferences {
  let meta = DocumentJsonMeta.fromMarkdown(tree);
  if (!meta) throw new Error("Missing YAML frontmatter");
  let references = new DocumentReferences(meta);

  visit(tree, "heading", (node) => {
    let reference = references.getSection(node.depth, node.data?.id);
    if (reference) {
      node.children.unshift({
        type: "text",
        value: reference + " ",
      });
    }
  });

  return references;
}

function parseAttributes(attributeString) {
  if (!attributeString.startsWith("{") || !attributeString.endsWith("}"))
    return;
  attributeString = attributeString.slice(1, -1);

  let regex =
    /(\w[\w0-9_-]*)=(?:(?:(?<quote>["'])((?:\\.|[^\\])*?)\k<quote>)|((?:[^'"]\S*)?))/g;

  let match;
  let result = Object.create(null);

  while ((match = regex.exec(attributeString)) !== null) {
    let key = match[1];
    let value;

    if (match[4]) {
      try {
        value = JSON.parse('"' + match[4] + '"');
      } catch (e) {
        value = match[4];
      }
    } else {
      value = match[3];
    }

    result[key] = value;
  }

  return result;
}

function parseImageAttrs(tree: mdast.Root) {
  visit(tree, (node) => {
    if ("children" in node) {
      for (let i = 1; i < node.children.length; i++) {
        let previous = node.children[i - 1];
        let current = node.children[i];

        if (previous.type !== "image" || current.type !== "text") {
          continue;
        }

        let text = current as mdast.Text;
        let image = previous as mdast.Image;

        let attrs = parseAttributes(text.value);
        if (attrs) {
          image.data = image.data || {};
          image.data = { attrs: attrs };
          node.children.splice(i, 1);
        }
      }
    }
  });
}

export interface ParsedMarkdown {
  tree: mdast.Root;
  references: DocumentReferences;
}
function defineElementReferences(src: string, references: DocumentReferences) {
  const declarations =
    /<div\s+class=["'](?:img|listing|table)-caption["']>([^\n]*?)<\/div>|\*\*(?:Определение|Утверждение|Лемма|Алгоритм)\s+(<span\s+class=(?:"ref"|ref)>[^<]+<\/span>)\*\*/g;
  for (const match of src.matchAll(declarations)) {
    const label = (match[1] ?? match[2]).match(
      /<span\s+class=(?:"ref"|ref)>([^<]+)<\/span>/,
    );
    if (!label) throw new Error("Caption has no label: " + match[0]);
    references.defineElement(label[1].trim());
  }
  for (const match of src.matchAll(
    /<span\s+class=(?:"(ref|cite)"|(ref|cite))>([^<]+)<\/span>/g,
  )) {
    const id = match[3].trim();
    if ((match[1] ?? match[2]) === "cite") references.getCite(id);
    else references.getReference(id);
  }
}

export function parseMarkdown(src: string): ParsedMarkdown {
  let tree = fromMarkdown(src, {
    extensions: [math(), gfmTable(), frontmatter(["yaml"])],
    mdastExtensions: [
      mathFromMarkdown(),
      gfmTableFromMarkdown(),
      frontmatterFromMarkdown(["yaml"]),
    ],
  });

  addHeaderReferences(tree);
  removePositions(tree);
  let references = addReferences(tree);
  defineElementReferences(src, references);
  parseImageAttrs(tree);

  return { tree, references };
}
