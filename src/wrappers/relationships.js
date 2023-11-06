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
class Relationships extends XML.Serializable {
    relations = new Map();
    readXml(xml) {
        this.relations = new Map();
        xml.getChild("Relationships")?.visitChildren("Relationship", (child) => {
            let id = child.getAttr("Id");
            let type = child.getAttr("Type");
            let target = child.getAttr("Target");
            if (id !== undefined && type !== undefined && target !== undefined) {
                this.relations.set(id, {
                    id: id,
                    type: type,
                    target: target
                });
            }
        });
        return this;
    }
    toXMLString() {
        let relations = Array.from(this.relations.values());
        let result = XML.Node.createDocument().appendChildren([
            XML.Node.build("Relationships")
                .setAttr("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
                .appendChildren(relations.map((relation) => {
                return XML.Node.build("Relationship")
                    .setAttr("Id", relation.id)
                    .setAttr("Type", relation.type)
                    .setAttr("Target", relation.target);
            }))
        ]);
        return result.toXmlString();
    }
    getUnusedId() {
        let prefix = "rId";
        let index = 1;
        while (this.relations.has(prefix + index)) {
            index++;
        }
        return prefix + index;
    }
    getRelForTarget(target) {
        for (let rel of this.relations.values()) {
            if (rel.target === target) {
                return rel;
            }
        }
    }
    join(other) {
        let result = new Map();
        for (let otherRel of other.relations.values()) {
            let existingTargetRel = this.getRelForTarget(otherRel.target);
            if (existingTargetRel) {
                result.set(otherRel.id, existingTargetRel.id);
                continue;
            }
            let newId = otherRel.id;
            if (this.relations.has(newId)) {
                newId = this.getUnusedId();
            }
            this.relations.set(newId, {
                ...otherRel,
                id: newId,
            });
            result.set(otherRel.id, newId);
        }
        return result;
    }
}
exports.default = Relationships;
//# sourceMappingURL=relationships.js.map