import mdast from "mdast";
import yaml from "yaml";

type DocumentJsonSection = Array<DocumentJsonSection> | Object | string;

function metaType(value: DocumentJsonSection) {
  if (Array.isArray(value)) return "array";
  return typeof value;
}

export class DocumentJsonMeta {
  section: DocumentJsonSection;
  path: string;

  constructor(section: DocumentJsonSection, path: string = "") {
    this.section = section;
    this.path = path;
  }

  getSection(path: string) {
    let any = this.getChild(path);
    return new DocumentJsonMeta(any, this.getAbsPath(path));
  }

  isArray() {
    if (this.section === undefined) {
      return false;
    } else return Array.isArray(this.section);
  }

  isMap() {
    if (this.section === undefined) {
      return false;
    } else return typeof this.section === "object" && !this.isArray();
  }

  asArray() {
    if (this.section === undefined) {
      this.reportNotExistError("", "array");
    } else if (!this.isArray()) {
      this.reportWrongTypeError("", "array", metaType(this.section));
    } else {
      return (this.section as Array<DocumentJsonSection>).map(
        (element, index) => {
          return new DocumentJsonMeta(element, this.getAbsPath(String(index)));
        },
      );
    }
  }

  getKeys() {
    if (!this.section) {
      this.reportNotExistError("", "object");
    } else if (!this.isMap()) {
      this.reportWrongTypeError("", "object", metaType(this.section));
    } else {
      return Object.keys(this.section);
    }
  }

  getString(path: string = ""): string {
    let child = this.getChild(path);
    if (!child) {
      this.reportNotExistError(path, "MetaInlines");
    } else if (typeof child !== "string") {
      this.reportWrongTypeError(path, "MetaInlines", metaType(child));
    } else {
      return child;
    }
  }

  private reportNotExistError(relPath: string, expected: string): never {
    let absPath = this.getAbsPath(relPath);
    throw new Error(
      "Failed to parse document metadata: expected to have " +
        expected +
        " at path " +
        absPath,
    );
  }

  private reportWrongTypeError(
    relPath: string,
    expected: string,
    actual: string,
  ): never {
    let absPath = this.getAbsPath(relPath);
    throw new Error(
      "Failed to parse document metadata: expected " +
        expected +
        " at path " +
        absPath +
        ", got " +
        actual +
        " instead",
    );
  }

  private getAbsPath(relPath: string) {
    if (this.path.length) {
      if (relPath.length) {
        return this.path + "." + relPath;
      }
      return this.path;
    }
    return relPath;
  }

  getChild(path: string): DocumentJsonSection | undefined {
    if (!path.length) return this.section;

    let result = this.section;

    for (let component of path.split(".")) {
      if (component === "__proto__") return undefined;
      if (!result) return undefined;

      if (Array.isArray(result)) {
        let index = Number.parseInt(component);
        if (!Number.isNaN(index)) {
          result = result[index];
        }
      } else if (typeof result === "object") {
        result = result[component];
      } else {
        return undefined;
      }
    }
    return result;
  }

  static fromMarkdown(markdown: mdast.Root) {
    let child = markdown.children[0];
    if (!child) return null;
    if (child.type !== "yaml") return null;

    return new DocumentJsonMeta(yaml.parse(child.value));
  }
}
