import {metaElementToSource, PandocMetaValue} from "src/pandoc/pandoc-json";

export class PandocJsonMeta {
    section: PandocMetaValue
    path: string

    constructor(section: PandocMetaValue, path: string = "") {
        this.section = section
        this.path = path
    }

    getSection(path: string) {
        let any = this.getChild(path)
        return new PandocJsonMeta(any, this.getAbsPath(path))
    }

    asArray() {
        if(this.section === undefined) {
            this.reportNotExistError("", "MetaList")
        } else if(this.section.t !== "MetaList") {
            this.reportWrongTypeError("", "MetaList", this.section.t)
        } else {
            return this.section.c.map((element, index) => {
                return new PandocJsonMeta(element, this.getAbsPath(String(index)))
            })
        }
    }

    getString(path: string = ""): string {
        let child = this.getChild(path)
        if(!child) {
            this.reportNotExistError(path, "MetaInlines")
        } else if(child.t !== "MetaInlines") {
            this.reportWrongTypeError(path, "MetaInlines", child.t)
        } else {
            return metaElementToSource(child.c)
        }
    }

    private reportNotExistError(relPath: string, expected: string): never {
        let absPath = this.getAbsPath(relPath)
        throw new Error("Failed to parse document metadata: expected to have " + expected + " at path " + absPath)
    }

    private reportWrongTypeError(relPath: string, expected: string, actual: string): never {
        let absPath = this.getAbsPath(relPath)
        throw new Error("Failed to parse document metadata: expected " + expected + " at path " + absPath + ", got " +
            actual + " instead")
    }

    private getAbsPath(relPath: string) {
        if (this.path.length) {
            if(relPath.length) {
                return this.path + "." + relPath
            }
            return this.path
        }
        return relPath
    }

    getChild(path: string): PandocMetaValue | undefined {
        if(!path.length) return this.section

        let result = this.section

        for (let component of path.split(".")) {
            // Be safe from prototype pollution
            if(component === "__proto__") return undefined
            if (!result) return undefined

            if(result.t === "MetaMap") {
                result = result.c[component]
            }

            if(result.t === "MetaList") {
                let index = Number.parseInt(component)
                if(!Number.isNaN(index)) {
                    result = result.c[index]
                }
            }
        }
        return result
    }
}