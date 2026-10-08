import test from "../test.js";
import assert from "assert";
import * as stylesReader from "../../lib/docx/styles-reader.js";
import * as documents from "../../lib/documents.js";
import {createFootnotesReader} from "../../lib/docx/notes-reader.js";
import {createBodyReader} from "../../lib/docx/body-reader.js";
import {Element as XmlElement} from "../../lib/xml/index.js";
test('ID and body of footnote are read', function() {
    var bodyReader = new createBodyReader({styles: stylesReader.defaultStyles});
    var footnoteBody = [new XmlElement("w:p", {}, [])];
    var footnotes = createFootnotesReader(bodyReader)(
        new XmlElement("w:footnotes", {}, [
            new XmlElement("w:footnote", {"w:id": "1"}, footnoteBody)
        ])
    );
    assert.equal(footnotes.value.length, 1);
    assert.deepEqual(footnotes.value[0].body, [new documents.Paragraph([])]);
    assert.deepEqual(footnotes.value[0].noteId, "1");
});

footnoteTypeIsIgnored('continuationSeparator');
footnoteTypeIsIgnored('separator');

function footnoteTypeIsIgnored(type) {
    test('footnotes of type ' + type + ' are ignored', function() {
        var footnotes = createFootnotesReader()(
            new XmlElement("w:footnotes", {}, [
                new XmlElement("w:footnote", {"w:id": "1", "w:type": type}, [])
            ])
        );
        assert.equal(footnotes.value.length, 0);
    });
}
