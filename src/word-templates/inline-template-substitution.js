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
class InlineTemplateSubstitution {
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
    replaceInlineTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(XML.keys.text, (textNode) => {
                    textNode.setText(textNode.getText().replace(this.template, this.replacement));
                });
            });
        }
    }
    removeParagraphsWithTemplate(body) {
        for (let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i]);
            let found = false;
            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(XML.keys.text, (textNode) => {
                    let text = textNode.getText();
                    if (text.indexOf(this.template) !== null) {
                        found = true;
                    }
                });
                return !found;
            });
            if (found) {
                body.removeChild([i]);
                i--;
            }
        }
    }
    performIn(body) {
        if (this.replacement === "@none") {
            this.removeParagraphsWithTemplate(body);
        }
        else {
            this.replaceInlineTemplate(body);
        }
    }
    perform() {
        let document = this.document;
        let documentBody = document.document.resource.toXml().getChild("w:document").getChild("w:body");
        this.performIn(documentBody);
        for (let header of document.headers) {
            this.performIn(header.resource.toXml().getChild("w:hdr"));
        }
        for (let footer of document.footers) {
            this.performIn(footer.resource.toXml().getChild("w:ftr"));
        }
    }
}
exports.default = InlineTemplateSubstitution;
//# sourceMappingURL=inline-template-substitution.js.map