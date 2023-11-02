"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DocumentMeta = exports.convertMetaToJsonRecursive = exports.getMetaString = exports.pandoc = void 0;
const child_process_1 = require("child_process");
function pandoc(src, args) {
    return new Promise((resolve, reject) => {
        let stdout = "";
        let stderr = "";
        let pandocProcess = (0, child_process_1.spawn)('pandoc', args);
        pandocProcess.stdin.end(src, 'utf-8');
        pandocProcess.stdout.on('data', (data) => {
            stdout += data;
        });
        pandocProcess.stderr.on('data', (data) => {
            stderr += data;
        });
        pandocProcess.on('exit', function (code) {
            if (stderr.length) {
                console.error("There was some pandoc warnings along the way:");
                console.error(stderr);
            }
            if (code == 0) {
                resolve(stdout);
            }
            else {
                reject(new Error("Pandoc returned non-zero exit code"));
            }
        });
    });
}
exports.pandoc = pandoc;
function getMetaString(value) {
    if (Array.isArray(value)) {
        let result = "";
        for (let component of value) {
            result += getMetaString(component);
        }
        return result;
    }
    if (typeof value !== "object" || !value.t) {
        return "";
    }
    if (value.t === "Str") {
        return value.c;
    }
    if (value.t === "Strong") {
        return "__" + getMetaString(value.c) + "__";
    }
    if (value.t === "Emph") {
        return "_" + getMetaString(value.c) + "_";
    }
    if (value.t === "Cite") {
        return getMetaString(value.c[1]);
    }
    if (value.t === "Space") {
        return " ";
    }
    if (value.t === "Link") {
        return getMetaString(value.c[1]);
    }
    return getMetaString(value.c);
}
exports.getMetaString = getMetaString;
function convertMetaToJsonRecursive(meta) {
    if (meta.t === "MetaList") {
        return meta.c.map((element) => {
            return convertMetaToJsonRecursive(element);
        });
    }
    if (meta.t === "MetaMap") {
        let result = {};
        for (let key of Object.getOwnPropertyNames(meta.c)) {
            // Be safe from prototype pollution
            if (key === "__proto__")
                continue;
            result[key] = convertMetaToJsonRecursive(meta.c[key]);
        }
        return result;
    }
    if (meta.t === "MetaInlines") {
        return getMetaString(meta.c);
    }
}
exports.convertMetaToJsonRecursive = convertMetaToJsonRecursive;
class DocumentMeta {
    section;
    path;
    constructor(section, path = "") {
        this.section = section;
        this.path = path;
    }
    getSection(path) {
        let any = this.getAny(path);
        return new DocumentMeta(any, this.getAbsPath(path));
    }
    asArray() {
        if (this.section === undefined) {
            this.reportNotExistError("", "array");
        }
        if (!Array.isArray(this.section)) {
            this.reportWrongTypeError("", "array", typeof this.section);
        }
        return this.section.map((element, index) => {
            return new DocumentMeta(element, this.getAbsPath(String(index)));
        });
    }
    getString(path = "") {
        let any = this.getAny(path);
        this.maybeReportError(path, "string", any);
        return any;
    }
    getNumber(path = "") {
        let any = this.getAny(path);
        this.maybeReportError(path, "number", any);
        return any;
    }
    maybeReportError(relPath, expectedType, got) {
        if (got === undefined) {
            this.reportNotExistError(relPath, expectedType);
        }
        if (typeof got !== expectedType) {
            let actualType = typeof got;
            this.reportWrongTypeError(relPath, expectedType, actualType);
        }
    }
    reportNotExistError(relPath, expected) {
        let absPath = this.getAbsPath(relPath);
        throw new Error("Failed to parse document metadata: expected to have " + expected + " at path " + absPath);
    }
    reportWrongTypeError(relPath, expected, actual) {
        let absPath = this.getAbsPath(relPath);
        throw new Error("Failed to parse document metadata: expected " + expected + " at path " + absPath + ", got " +
            actual + " instead");
    }
    getAbsPath(relPath) {
        if (this.path.length) {
            if (relPath.length) {
                return this.path + "." + relPath;
            }
            return this.path;
        }
        return relPath;
    }
    getAny(path) {
        if (!path.length)
            return this.section;
        let result = this.section;
        for (let component of path.split(".")) {
            // Be safe from prototype pollution
            if (component === "__proto__")
                return undefined;
            if (!result)
                return undefined;
            result = result[component];
        }
        return result;
    }
    static fromPandocMeta(meta) {
        let result = {};
        for (let key of Object.getOwnPropertyNames(meta)) {
            // Be safe from prototype pollution
            if (key === "__proto__")
                continue;
            result[key] = convertMetaToJsonRecursive(meta[key]);
        }
        return new DocumentMeta(result);
    }
}
exports.DocumentMeta = DocumentMeta;
//# sourceMappingURL=pandoc.js.map