import WordDocument from "src/word/word-document";
import * as XML from "src/xml";

export default class ParagraphTemplateSubstitution {
    private document: WordDocument;
    private template: string;
    private replacement: () => XML.Node[];

    setDocument(document: WordDocument) {
        this.document = document
        return this
    }

    setTemplate(template: string) {
        this.template = template
        return this
    }

    setReplacement(replacement: () => XML.Node[]) {
        this.replacement = replacement
        return this
    }

    perform() {
        const body = this.document.document.resource.toXml().getChild("w:document").getChild("w:body")
        this.replaceParagraphsWithTemplate(body)
        return this
    }

    private replaceParagraphsWithTemplate(body: XML.Node) {
        for(let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i])

            let paragraphText = ""
            child.visitSubtree("w:t", (textNode) => {
                paragraphText += XML.getTextContents(textNode)
            })

            if(paragraphText.indexOf(this.template) === -1) {
                continue
            }

            if(paragraphText !== this.template) {
                throw new Error(`The ${this.template} pattern should be the only text of the paragraph`)
            }

            body.removeChild([i])
            let replacement = this.replacement()
            body.insertChildren(replacement, [i])
            i += replacement.length - 1
        }
    }
}