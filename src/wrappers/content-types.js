"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const XML = __importStar(require("../xml.js"));
class ContentTypes extends XML.Serializable {
    defaults;
    overrides;
    readXml(xml) {
        this.defaults = [];
        this.overrides = [];
        let types = xml.getChild("Types");
        types?.visitChildren("Default", (child) => {
            let extension = child.getAttr("Extension");
            let contentType = child.getAttr("ContentType");
            if (extension !== undefined && contentType !== undefined) {
                this.defaults.push({
                    extension: extension,
                    contentType: contentType
                });
            }
        });
        types?.visitChildren("Override", (child) => {
            let partName = child.getAttr("PartName");
            let contentType = child.getAttr("ContentType");
            if (partName !== undefined && contentType !== undefined) {
                this.overrides.push({
                    partName: partName,
                    contentType: contentType
                });
            }
        });
        return this;
    }
    toXml() {
        return XML.Node.createDocument().appendChildren([
            XML.Node.build("Types")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types")
                .appendChildren(this.defaults.map((def) => {
                return XML.Node.build("Default")
                    .setAttr("Extension", def.extension)
                    .setAttr("ContentType", def.contentType);
            }))
                .appendChildren(this.overrides.map((override) => {
                return XML.Node.build("Override")
                    .setAttr("PartName", override.partName)
                    .setAttr("ContentType", override.contentType);
            }))
        ]);
    }
    getContentTypeForExt(ext) {
        for (let def of this.defaults) {
            if (def.extension === ext)
                return def.contentType;
        }
        return null;
    }
    getOverrideForPartName(partName) {
        for (let override of this.overrides) {
            if (override.partName === partName)
                return override.contentType;
        }
        return null;
    }
    join(other) {
        for (let otherDef of other.defaults) {
            if (this.getContentTypeForExt(otherDef.extension) === null) {
                this.defaults.push({
                    ...otherDef
                });
            }
        }
        for (let otherOverride of other.overrides) {
            if (this.getOverrideForPartName(otherOverride.partName) === null) {
                this.overrides.push({
                    ...otherOverride
                });
            }
        }
    }
}
exports.default = ContentTypes;
//# sourceMappingURL=content-types.js.map