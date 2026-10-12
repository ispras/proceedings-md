import { execFileSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

export function exportPdf(docx: string) {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "proceedings-pdf-"));
  const expected = path.join(
    temporary,
    path.basename(docx).replace(/\.docx$/i, ".pdf"),
  );
  const target = path.resolve(docx).replace(/\.docx$/i, ".pdf");
  if (target === path.resolve(docx))
    throw new Error("PDF export requires a .docx output path");
  try {
    execFileSync(
      "soffice",
      [
        `-env:UserInstallation=${pathToFileURL(path.join(temporary, "profile")).href}`,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        temporary,
        path.resolve(docx),
      ],
      { stdio: "pipe", timeout: 120_000 },
    );
    if (!fs.existsSync(expected) || fs.statSync(expected).size === 0)
      throw new Error("LibreOffice produced no PDF");
    fs.copyFileSync(expected, target);
    console.error(`Built ${target}`);
  } catch (error: any) {
    throw new Error(
      `PDF export failed: ${error.stderr?.toString() || error.message}`,
    );
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
}
