import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { PNG } from "pngjs";
import { parseMarkdown } from "../tools/markdown/markdown";
import { build } from "../tools/build";
import { validateDocx } from "../tools/validate";
import { figureDependencies } from "../tools/figures";
import { markerDepth } from "../tools/markdown/formula-render";

const root = fileURLToPath(new URL("..", import.meta.url));
const article = fs.readFileSync(path.join(root, "sample/sample.md"), "utf8");
const frontmatter = article.slice(0, article.indexOf("\n---", 4) + 4);

function parse(body: string) {
  return parseMarkdown(`${frontmatter}\n\n${body}`);
}

test("forward references follow definition order for figures and theorems", () => {
  const { references } = parse(
    `# Intro {#intro}\n\nSee <span class=ref>fig:b</span>, <span class=ref>lem:b</span>.\n\n<div class="img-caption">Fig. <span class=ref>fig:a</span>. A</div>\n\n<div class="img-caption">Fig. <span class=ref>fig:b</span>. B</div>\n\n**Лемма <span class=ref>lem:a</span>**. A\n\n**Лемма <span class=ref>lem:b</span>**. B`,
  );
  assert.equal(references.getReference("fig:b"), "2");
  assert.equal(references.getReference("lem:b"), "2");
  assert.equal(references.getReference("intro"), "1");
});

test("missing and duplicate references fail instead of inventing numbers", () => {
  assert.throws(
    () => parse("See <span class=ref>fig:missing</span>."),
    /Undefined reference/,
  );
  assert.throws(
    () => parse("<span class=cite>missing</span>"),
    /Undefined citation/,
  );
  assert.throws(
    () => parse("# One {#same}\n\n# Two {#same}"),
    /Duplicate label/,
  );
  assert.throws(
    () =>
      parse(
        '<div class="img-caption">Fig. <span class=ref>fig:a</span></div>\n\n<div class="img-caption">Fig. <span class=ref>fig:a</span></div>',
      ),
    /Duplicate label/,
  );
  assert.throws(() => parse("# One\n\n### Skipped"), /Heading skips/);
});

test("baseline detection preserves descenders and rejects a missing marker", () => {
  const clean = new PNG({ width: 8, height: 12 });
  const marked = new PNG({ width: 8, height: 12 });
  marked.data.fill(0);
  const offset = (7 * 8 + 2) * 4;
  marked.data[offset] = 255;
  marked.data[offset + 3] = 255;
  assert.equal(
    markerDepth(PNG.sync.write(clean), PNG.sync.write(marked)),
    (4 / 600) * 72,
  );
  assert.throws(
    () => markerDepth(PNG.sync.write(clean), PNG.sync.write(clean)),
    /marker is missing/,
  );
});

test("DOCX builds from another directory, preserves formatting and has stable bytes", async () => {
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), "proceedings-test-"));
  try {
    fs.mkdirSync(path.join(temp, "figures"));
    for (const ext of ["png", "svg"])
      fs.copyFileSync(
        path.join(root, `test/fixtures/image.${ext}`),
        path.join(temp, `figures/image.${ext}`),
      );
    const input = path.join(temp, "article.md");
    fs.writeFileSync(
      input,
      `${frontmatter}\n\n# Example {#example}\n\nA **bold** and *italic* word, x<sub>i</sub>, x<sup>2</sup> and \`code\`.\n\n> Quoted text survives.\n\nFirst line  \nSecond line\n\n![](figures/image.png)\n\n<div class="img-caption">Fig. <span class=ref>fig:image</span>. Image</div>\n\n| A | B |\n|:-|--:|\n| 1 | 2 |`,
    );
    const output = path.join(temp, "out", "article.docx");
    await build(input, output);
    const first = fs.readFileSync(output);
    await build(input, output);
    assert.deepEqual(fs.readFileSync(output), first);
    const zip = await JSZip.loadAsync(first);
    const xml = await zip.file("word/document.xml")!.async("string");
    assert.match(
      xml.replace(/<([\w:]+)([^>]*)><\/\1>/g, "<$1$2/>"),
      /<w:pPr><w:pStyle w:val="688"\/><w:numPr><w:ilvl w:val="0"\/><w:numId w:val="0"\/><\/w:numPr><\/w:pPr>/,
    );
    assert.match(xml, /Quoted text survives/);
    assert.match(xml, /ORCID: 0000-0000-0000-0000/);
    assert.match(xml, /w:highlight w:val="yellow"/);
    assert.match(xml, /w:vertAlign w:val="subscript"/);
    assert.match(xml, /w:vertAlign w:val="superscript"/);
    assert.match(xml, /<w:br(?:\/>|>)/);
    assert.match(xml, /<w:tblPr>/);
    assert.match(xml, /asvg:svgBlip/);
    assert.doesNotMatch(xml, /TIKZ::|\{\{\{/);
    zip.file("word/media/broken.png", "broken");
    zip.file(
      "word/_rels/document.xml.rels",
      (await zip.file("word/_rels/document.xml.rels")!.async("string")).replace(
        /Target="media\/[^"]+"/,
        'Target="media/missing.png"',
      ),
    );
    const broken = path.join(temp, "broken.docx");
    fs.writeFileSync(broken, await zip.generateAsync({ type: "nodebuffer" }));
    await assert.rejects(validateDocx(broken), /Broken relationship/);
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
});

test("missing input assets fail before rendering", async () => {
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), "proceedings-test-"));
  try {
    const input = path.join(temp, "article.md");
    fs.writeFileSync(input, `${frontmatter}\n\n![](missing.png)`);
    await assert.rejects(
      build(input, path.join(temp, "out.docx"), true),
      /Missing image/,
    );
    assert.deepEqual(fs.readdirSync(temp), ["article.md"]);
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
});

test("figure dependencies are collected and cycles rejected", () => {
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), "proceedings-test-"));
  try {
    fs.writeFileSync(
      path.join(temp, "root.tex"),
      String.raw`\input{left.tex}\input{right.tex}`,
    );
    fs.writeFileSync(path.join(temp, "left.tex"), "Left");
    fs.writeFileSync(path.join(temp, "right.tex"), "Right");
    assert.equal(figureDependencies(path.join(temp, "root.tex")).size, 3);
    const source = path.join(temp, "cycle.tex");
    fs.writeFileSync(source, String.raw`\input{cycle.tex}`);
    assert.throws(() => figureDependencies(source), /Cyclic figure input/);
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
});
