import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { visit } from "unist-util-visit";
import { parseMarkdown } from "./markdown/markdown";
import { renderFigures } from "./figures";

try {
  const input = path.resolve(
    process.argv[2] ??
      fileURLToPath(new URL("../sample/sample.md", import.meta.url)),
  );
  const { tree } = parseMarkdown(fs.readFileSync(input, "utf8"));
  const urls: string[] = [];
  visit(tree, "image", (node) => urls.push(node.url));
  const paths = renderFigures(path.dirname(input), urls);
  for (const [target, rendered] of paths) {
    fs.copyFileSync(rendered, target);
    fs.copyFileSync(
      rendered.replace(/\.png$/, ".svg"),
      target.replace(/\.png$/, ".svg"),
    );
  }
  console.error(`Updated ${paths.size} PNG/SVG pairs`);
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
}
