import { DocumentJsonMeta } from "../document-json-meta";

export default class DocumentReferences {
  private stack: number[] = [];
  private labels = new Map<string, string>();
  private counts = new Map<string, number>();
  private citations = new Map<string, string>();

  constructor(meta: DocumentJsonMeta) {
    const links = meta.getSection("ispras_templates.links").asArray();
    links.forEach((link, index) => {
      const id = link.getString("id");
      if (this.citations.has(id))
        throw new Error(`Duplicate bibliography ID: ${id}`);
      this.citations.set(id, `[${index + 1}]`);
    });
  }

  private define(id: string, number: string) {
    if (this.labels.has(id)) throw new Error(`Duplicate label: ${id}`);
    this.labels.set(id, number);
  }

  getSection(depth: number, id?: string): string | undefined {
    if (depth > 3) return;
    if (depth > this.stack.length + 1)
      throw new Error(`Heading skips a level: ${depth}`);
    if (depth <= this.stack.length) {
      this.stack.length = depth;
      this.stack[depth - 1]++;
    } else this.stack.push(1);
    const number = this.stack.join(".");
    if (id) this.define(id, number);
    return number + (depth === 1 ? "." : "");
  }

  defineElement(id: string) {
    const prefix = id.split(":")[0];
    if (!["fig", "lst", "tab", "def", "stmt", "lem", "alg"].includes(prefix)) {
      throw new Error(`Unsupported element label: ${id}`);
    }
    const count = (this.counts.get(prefix) ?? 0) + 1;
    this.counts.set(prefix, count);
    this.define(id, String(count));
  }

  getReference(id: string): string {
    const number = this.labels.get(id);
    if (number === undefined) throw new Error(`Undefined reference: ${id}`);
    return number;
  }

  getCite(id: string): string {
    const number = this.citations.get(id);
    if (number === undefined) throw new Error(`Undefined citation: ${id}`);
    return number;
  }
}
