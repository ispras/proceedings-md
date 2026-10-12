import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { latexPreamble, runTool, checkLatexLog } from "./latex";

export function figureDependencies(
  source: string,
  chain: string[] = [],
): Map<string, string> {
  if (chain.includes(source)) throw new Error(`Cyclic figure input: ${source}`);
  const content = fs.readFileSync(source, "utf8");
  const result = new Map([[source, content]]);
  for (const match of content.matchAll(/\\input\{([^}]+)\}/g)) {
    const input = path.resolve(
      path.dirname(source),
      match[1].endsWith(".tex") ? match[1] : match[1] + ".tex",
    );
    for (const [filename, text] of figureDependencies(input, [
      ...chain,
      source,
    ]))
      result.set(filename, text);
  }
  return result;
}

const figurePackages = String.raw`
\usetikzlibrary{patterns,patterns.meta,arrows.meta,positioning,shapes,shapes.geometric,fit,backgrounds,matrix,decorations.pathreplacing}
\usepackage{pgfplots}
\usepgfplotslibrary{groupplots}
\pgfplotsset{compat=1.18}
\usepackage{pgf-pie}
\setlength\PreviewBorder{2pt}
`;

export function renderFigures(
  sourceDir: string,
  urls: string[],
): Map<string, string> {
  const preamble =
    latexPreamble().replace(
      "active,tightpage,displaymath,textmath",
      "active,tightpage",
    ) + figurePackages;
  const engine = runTool("xelatex", ["--version"], os.tmpdir());
  const paths = new Map<string, string>();
  const missing: { input: string; output: string; source: string }[] = [];
  for (const url of new Set(urls)) {
    const input = path.resolve(sourceDir, url);
    const stem = path.basename(input, path.extname(input));
    const source = path.join(sourceDir, "figures", "source", stem + ".tex");
    if (!fs.existsSync(source))
      throw new Error(`Missing figure source: ${source}`);
    const content = [...figureDependencies(source)].map(([filename, text]) => [
      path.basename(filename),
      text,
    ]);
    const key = createHash("sha256")
      .update(JSON.stringify([preamble, engine, content, 300, 1]))
      .digest("hex");
    const output = path.join(
      sourceDir,
      ".cache",
      "figures",
      key,
      stem + ".png",
    );
    paths.set(input, output);
    if (
      !fs.existsSync(output) ||
      !fs.existsSync(output.replace(/\.png$/, ".svg"))
    )
      missing.push({ input, output, source });
  }
  console.error(`Figures: ${paths.size} unique, ${missing.length} to render`);
  if (!missing.length) return paths;
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), "proceedings-figures-"));
  try {
    let tex = preamble + "\n\\begin{document}\n";
    missing.forEach((figure, index) => {
      for (const filename of figureDependencies(figure.source).keys())
        fs.copyFileSync(filename, path.join(temp, path.basename(filename)));
      fs.copyFileSync(figure.source, path.join(temp, `figure-${index}.tex`));
      tex += `\\begin{preview}\\input{figure-${index}.tex}\\end{preview}\n`;
    });
    tex += "\\end{document}\n";
    fs.writeFileSync(path.join(temp, "figures.tex"), tex);
    runTool(
      "xelatex",
      [
        "-no-shell-escape",
        "-halt-on-error",
        "-interaction=nonstopmode",
        "figures.tex",
      ],
      temp,
    );
    checkLatexLog(fs.readFileSync(path.join(temp, "figures.log"), "utf8"));
    runTool(
      "pdftocairo",
      ["-png", "-transp", "-r", "300", "figures.pdf", "page"],
      temp,
    );
    const pages = new Map<number, string>();
    for (const file of fs.readdirSync(temp)) {
      const match = file.match(/^page-(\d+)\.png$/);
      if (match) pages.set(Number(match[1]), file);
    }
    if (pages.size !== missing.length)
      throw new Error(
        `Expected ${missing.length} figure pages, got ${pages.size}`,
      );
    missing.forEach((figure, index) => {
      fs.mkdirSync(path.dirname(figure.output), { recursive: true });
      const svg = path.join(temp, `figure-${index}.svg`);
      // Both formats come from the same PDF page, including its clipping bounds.
      runTool(
        "pdftocairo",
        [
          "-svg",
          "-f",
          String(index + 1),
          "-l",
          String(index + 1),
          "figures.pdf",
          svg,
        ],
        temp,
      );
      fs.copyFileSync(svg, figure.output.replace(/\.png$/, ".svg"));
      fs.copyFileSync(path.join(temp, pages.get(index + 1)!), figure.output);
    });
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
  return paths;
}
