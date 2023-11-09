import WordDocument from "src/word/word-document";
import * as XML from "src/xml";

export default class InlineTemplateSubstitution {
    private document: WordDocument;
    private template: string;
    private replacement: string;

    setDocument(document: WordDocument) {
        this.document = document
        return this
    }

    setTemplate(template: string) {
        this.template = template
        return this
    }

    setReplacement(replacement: string) {
        this.replacement = replacement
        return this
    }

    private replaceInlineTemplate(body: XML.Node) {
        for(let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i])

            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(XML.keys.text, (textNode) => {
                    textNode.setText(textNode.getText().replace(this.template, this.replacement))
                })
            })
        }
    }

    private removeParagraphsWithTemplate(body: XML.Node) {
        for(let i = 0; i < body.getChildrenCount(); i++) {
            let child = body.getChild([i])
            let found = false

            child.visitSubtree("w:t", (paragraphText) => {
                paragraphText.visitSubtree(XML.keys.text, (textNode) => {
                    let text = textNode.getText()

                    if(text.indexOf(this.template) !== null) {
                        found = true
                    }
                })

                return !found
            })

            if(found) {
                body.removeChild([i])
                i--
            }
        }
    }

    performIn(body: XML.Node) {
        if (this.replacement === "@none") {
            this.removeParagraphsWithTemplate(body)
        } else {
            this.replaceInlineTemplate(body)
        }
    }

    perform() {
        let document = this.document

        let documentBody = document.document.resource.toXml().getChild("w:document").getChild("w:body")
        this.performIn(documentBody)

        for(let header of document.headers) {
            this.performIn(header.resource.toXml().getChild("w:hdr"))
        }

        for(let footer of document.footers) {
            this.performIn(footer.resource.toXml().getChild("w:ftr"))
        }

        return this
    }
}
