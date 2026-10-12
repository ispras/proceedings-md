import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { PNG } from "pngjs";
import { latexPreamble, runTool as run, checkLatexLog } from "../latex";

export interface FormulaMetrics {
  formula: string;
  width: number;
  height: number;
  heightAbove: number;
  depth: number;
  png: Buffer;
  svg: Buffer | null;
}
export interface FormulaRequest {
  formula: string;
  display: boolean;
}
export function formulaKey(formula: string, display: boolean) {
  return (display ? "D:" : "I:") + formula;
}
const DPI = 600;
// A separate marked copy locates the baseline without adding pixels to the image.
export function markerDepth(clean: Buffer, marked: Buffer): number {
  const image = PNG.sync.read(marked);
  const height = PNG.sync.read(clean).height;
  for (let y = image.height - 1; y >= 0; y--) {
    for (let x = 0; x < image.width; x++) {
      const i = (y * image.width + x) * 4;
      if (
        image.data[i] > 140 &&
        image.data[i + 1] < 115 &&
        image.data[i + 2] < 115 &&
        image.data[i + 3] > 127
      ) {
        return (Math.max(0, height - 1 - y) / DPI) * 72;
      }
    }
  }
  throw new Error("Formula baseline marker is missing");
}

export function renderFormulas(
  requests: FormulaRequest[],
  cacheDir: string,
): Map<string, FormulaMetrics> {
  const result = new Map<string, FormulaMetrics>();
  if (!requests.length) return result;
  const preamble = latexPreamble();
  const engine = run("xelatex", ["--version"], os.tmpdir());
  const unique = [
    ...new Map(
      requests.map((r) => [formulaKey(r.formula, r.display), r]),
    ).values(),
  ];
  const hashes = new Map<string, string>();
  const missing: FormulaRequest[] = [];
  for (const request of unique) {
    const key = formulaKey(request.formula, request.display);
    const hash = createHash("sha256")
      .update(JSON.stringify([preamble, engine, DPI, key, 1]))
      .digest("hex");
    hashes.set(key, hash);
    const prefix = path.join(cacheDir, hash);
    if (fs.existsSync(prefix + ".json") && fs.existsSync(prefix + ".png")) {
      const metrics = JSON.parse(fs.readFileSync(prefix + ".json", "utf8"));
      result.set(key, {
        ...metrics,
        png: fs.readFileSync(prefix + ".png"),
        svg: null,
      });
    } else missing.push(request);
  }
  console.error(
    `Formulas: ${unique.length} unique, ${missing.length} to render`,
  );
  if (!missing.length) return result;
  fs.mkdirSync(cacheDir, { recursive: true });
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), "proceedings-formulas-"));
  try {
    let tex = preamble + "\n\\begin{document}\n";
    for (const request of missing) {
      const open = request.display ? "\\[" : "$";
      const close = request.display ? "\\]" : "$";
      tex += `\\begin{preview}${open}${request.formula}${close}\\end{preview}\n`;
      tex += `\\begin{preview}${open}\\sbbaselinemark{}${request.formula}${close}\\end{preview}\n`;
    }
    tex += "\\end{document}\n";
    fs.writeFileSync(path.join(temp, "formulas.tex"), tex);
    run(
      "xelatex",
      [
        "-no-shell-escape",
        "-halt-on-error",
        "-interaction=nonstopmode",
        "formulas.tex",
      ],
      temp,
    );
    const log = fs.readFileSync(path.join(temp, "formulas.log"), "utf8");
    checkLatexLog(log);
    run(
      "pdftocairo",
      ["-png", "-transp", "-r", String(DPI), "formulas.pdf", "page"],
      temp,
    );
    const pages = new Map<number, string>();
    for (const name of fs.readdirSync(temp)) {
      const match = name.match(/^page-(\d+)\.png$/);
      if (match) pages.set(Number(match[1]), name);
    }
    if (pages.size !== missing.length * 2)
      throw new Error(
        `Expected ${missing.length * 2} formula pages, got ${pages.size}`,
      );
    missing.forEach((request, index) => {
      const png = fs.readFileSync(path.join(temp, pages.get(2 * index + 1)!));
      const marked = fs.readFileSync(
        path.join(temp, pages.get(2 * index + 2)!),
      );
      const dimensions = PNG.sync.read(png);
      const width = (dimensions.width / DPI) * 72;
      const height = (dimensions.height / DPI) * 72;
      const depth = markerDepth(png, marked);
      const metrics = {
        formula: request.formula,
        width,
        height,
        depth,
        heightAbove: height - depth,
      };
      const key = formulaKey(request.formula, request.display);
      const prefix = path.join(cacheDir, hashes.get(key)!);
      fs.writeFileSync(prefix + ".png", png);
      fs.writeFileSync(prefix + ".json", JSON.stringify(metrics));
      result.set(key, { ...metrics, png, svg: null });
    });
  } finally {
    fs.rmSync(temp, { recursive: true, force: true });
  }
  return result;
}
