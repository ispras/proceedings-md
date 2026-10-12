import fs from "node:fs";
import path from "node:path";
import { visit } from "unist-util-visit";
import type { Root } from "mdast";
import { DocumentJsonMeta } from "./document-json-meta";
import JSZip from "jszip";
import { XMLValidator } from "fast-xml-parser";
import * as XML from "./xml";

export function validateSource(
  tree: Root,
  sourceDir: string,
  rebuildFigures = false,
) {
  const meta = DocumentJsonMeta.fromMarkdown(tree);
  if (!meta?.getSection("ispras_templates").isMap())
    throw new Error("Missing ispras_templates metadata");
  const section = meta.getSection("ispras_templates");
  for (const language of ["ru", "en"]) {
    for (const field of [
      "header",
      "abstract",
      "keywords",
      "for_citation",
      "page_header",
      "acknowledgements",
    ]) {
      section.getString(`${field}_${language}`);
    }
  }
  const organizations = new Set(
    section
      .getSection("organizations")
      .asArray()
      .map((org) => {
        org.getString("name_ru");
        org.getString("name_en");
        return org.getString("id");
      }),
  );
  for (const author of section.getSection("authors").asArray()) {
    for (const field of [
      "name_ru",
      "name_en",
      "orcid",
      "email",
      "details_ru",
      "details_en",
    ])
      author.getString(field);
    for (const org of author.getSection("organizations").asArray()) {
      if (!organizations.has(org.getString()))
        throw new Error(`Unknown author organization: ${org.getString()}`);
    }
  }
  section
    .getSection("links")
    .asArray()
    .forEach((link) => link.getString("description"));
  visit(tree, "image", (node) => {
    const filename = path.resolve(sourceDir, node.url);
    const expected = rebuildFigures
      ? path.join(
          sourceDir,
          "figures",
          "source",
          path.basename(filename, path.extname(filename)) + ".tex",
        )
      : filename;
    if (!fs.existsSync(expected)) throw new Error(`Missing image: ${expected}`);
  });
  return section;
}

export async function validateDocx(filename: string) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filename));
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || !/\.(xml|rels)$/.test(entry.name)) continue;
    const text = await entry.async("string");
    const valid = XMLValidator.validate(text);
    if (valid !== true) throw new Error(`${entry.name}: ${valid.err.msg}`);
    if (text.includes("{{{"))
      throw new Error(`Unresolved template in ${entry.name}`);
    const document = XML.Node.fromXmlString(text);
    if (entry.name.endsWith(".rels")) {
      const ids = new Set<string>();
      document
        .getChild("Relationships")
        ?.visitChildren("Relationship", (node) => {
          const id = node.getAttr("Id")!;
          if (ids.has(id)) throw new Error(`Duplicate relationship ${id}`);
          ids.add(id);
          if (node.getAttr("TargetMode") === "External") return;
          const target = node.getAttr("Target")!;
          const base =
            entry.name === "_rels/.rels"
              ? ""
              : path.posix.dirname(path.posix.dirname(entry.name));
          const resolved = target.startsWith("/")
            ? target.slice(1)
            : path.posix.normalize(path.posix.join(base, target));
          if (!zip.file(resolved))
            throw new Error(
              `Broken relationship: ${entry.name} -> ${resolved}`,
            );
        });
    }
    if (entry.name === "word/document.xml") {
      const ids = new Set<string>();
      document.visitSubtree("w:p", (p) => {
        p.visitSubtree("w:p", (nested) => {
          throw new Error("Nested Word paragraph");
        });
      });
      document.visitSubtree("wp:docPr", (node) => {
        const id = node.getAttr("id")!;
        if (!/^\d+$/.test(id) || ids.has(id))
          throw new Error(`Invalid drawing ID: ${id}`);
        ids.add(id);
      });
    }
  }
}
