import fs from "node:fs";
import { renderFigures } from "./figures";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { visit } from "unist-util-visit";
import { parseMarkdown } from "./markdown/markdown";
import { validateSource, validateDocx } from "./validate";
import WordDocument from "./word/word-document";
import {
  generateDocxBody,
  normalizeFormula,
  substituteTemplates,
} from "./generator/generator";
import { renderFormulas, type FormulaRequest } from "./markdown/formula-render";

export async function build(
  input: string,
  output: string,
  checkOnly = false,
  rebuildFigures = false,
) {
  const source = path.resolve(input);
  if (
    !checkOnly &&
    (!/\.docx$/i.test(output) || path.resolve(output) === source)
  )
    throw new Error("Output must be a .docx file distinct from the source");
  const sourceDir = path.dirname(source);
  const { tree, references } = parseMarkdown(fs.readFileSync(source, "utf8"));
  const meta = validateSource(tree, sourceDir, rebuildFigures);
  if (checkOnly) return;
  const requests: FormulaRequest[] = [];
  visit(tree, (node) => {
    if (node.type === "inlineMath" || node.type === "math")
      requests.push({
        formula: normalizeFormula(node.value),
        display: node.type === "math",
      });
  });
  const figureUrls: string[] = [];
  visit(tree, "image", (node) => figureUrls.push(node.url));
  const imagePaths = rebuildFigures
    ? renderFigures(sourceDir, figureUrls)
    : new Map<string, string>();
  const metrics = renderFormulas(
    requests,
    path.join(sourceDir, ".cache", "formulas"),
  );
  const template = fileURLToPath(
    new URL("../resources/isp-reference.docx", import.meta.url),
  );
  const doc = await new WordDocument().load(template);
  generateDocxBody(tree, doc, meta, references, metrics, sourceDir, imagePaths);
  substituteTemplates(doc, meta);
  const target = path.resolve(output);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const temporary = target + `.tmp-${process.pid}`;
  try {
    await doc.save(temporary);
    await validateDocx(temporary);
    fs.renameSync(temporary, target);
  } finally {
    fs.rmSync(temporary, { force: true });
  }
}
