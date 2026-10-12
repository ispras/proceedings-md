import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import fs from "node:fs";

export function latexPreamble() {
  const preamblePath = fileURLToPath(
    new URL("../resources/formulas.tex", import.meta.url),
  );
  const fontsDir = fileURLToPath(
    new URL("../resources/fonts/", import.meta.url),
  );
  return fs
    .readFileSync(preamblePath, "utf8")
    .replace("FONT_DIRECTORY", fontsDir.replaceAll("\\", "/"));
}

export function runTool(command: string, args: string[], cwd: string) {
  try {
    return execFileSync(command, args, {
      cwd,
      env: { ...process.env, XDG_CACHE_HOME: `${cwd}/.font-cache` },
      stdio: "pipe",
      encoding: "utf8",
      timeout: 120_000,
      maxBuffer: 16 * 1024 * 1024,
    });
  } catch (error: any) {
    const detail =
      error.stdout?.toString() || error.stderr?.toString() || error.message;
    throw new Error(`${command} failed:\n${detail.slice(-6000)}`);
  }
}

export function checkLatexLog(log: string) {
  const missing = log
    .split("\n")
    .filter((line) => line.includes("Missing character:"));
  if (missing.length)
    throw new Error("Font lacks glyphs:\n" + missing.join("\n"));
}
