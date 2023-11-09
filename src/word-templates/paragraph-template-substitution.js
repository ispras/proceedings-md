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
const XML = __importStar(require("src/xml"));
class ParagraphTemplateSubstitution {
    document;
    template;
    replacement;
    setDocument(document) {
        this.document = document;
        return this;
    }
    setTemplate(template) {
        this.template = template;
        return this;
    }
    setReplacement(replacement) {
        this.replacement = replacement;
        return this;
    }
    perform() {
        const body = this.document.document.resource.toXml().getChild("w:document").getChild("w:body");
        this.replaceParagraphsWithTemplate(body);
        return this;
    }
    replaceParagraphsWithTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            let paragraphText = "";
            child.visitSubtree("w:t", (textNode) => {
                paragraphText += XML.getTextContents(textNode);
            });
            if (paragraphText.indexOf(this.template) === -1) {
                continue;
            }
            if (paragraphText !== this.template) {
                throw new Error(`The ${this.template} pattern should be the only text of the paragraph`);
            }
            body.removeChild([i]);
            let replacement = this.replacement();
            body.insertChildren(replacement, [i]);
            i += replacement.length - 1;
        }
    }
}
exports.default = ParagraphTemplateSubstitution;
//# sourceMappingURL=paragraph-template-substitution.js.map