import { parseArgs } from "node:util";
import { fileURLToPath } from "node:url";
import { build } from "./build";
import { exportPdf } from "./pdf";

try {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      check: { type: "boolean" },
      pdf: { type: "boolean" },
      figures: { type: "boolean" },
      help: { type: "boolean", short: "h" },
    },
  });
  if (values.help)
    console.log(
      "Usage: npm run build -- [--check] [--pdf] [--figures] [input.md [output.docx]]",
    );
  else {
    if (positionals.length > 2)
      throw new Error("Expected at most input and output paths");
    const input =
      positionals[0] ??
      fileURLToPath(new URL("../sample/sample.md", import.meta.url));
    const output =
      positionals[1] ??
      fileURLToPath(new URL("../build/sample.docx", import.meta.url));
    if (values.check && values.pdf)
      throw new Error("--check and --pdf cannot be combined");
    await build(input, output, values.check, values.figures);
    if (values.pdf) exportPdf(output);
    console.error(values.check ? "Source checks passed" : `Built ${output}`);
  }
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
}
