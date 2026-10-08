import test from "../test.js";
import assert from "assert";
import * as documents from "../../lib/documents.js";
import * as xml from "../../lib/xml/index.js";
import {DocumentXmlReader} from "../../lib/docx/document-xml-reader.js";
import {createBodyReaderForTests} from "./testing.js";
test("when body element is present then body is read", function() {
    var bodyReader = createBodyReaderForTests({});
    var documentXmlReader = new DocumentXmlReader({
        bodyReader: bodyReader
    });
    var textXml = xml.element("w:t", {}, [xml.text("Hello!")]);
    var runXml = xml.element("w:r", {}, [textXml]);
    var paragraphXml = xml.element("w:p", {}, [runXml]);
    var bodyXml = xml.element("w:body", {}, [paragraphXml]);
    var documentXml = xml.element("w:document", {}, [bodyXml]);

    var result = documentXmlReader.convertXmlToDocument(documentXml);

    assert.deepEqual(result.messages, []);
    assert.deepEqual(result.value, documents.document(
        [documents.paragraph([documents.run([documents.text("Hello!")])])],
        {}
    ));
});

test("when body element is not present then error is thrown", function() {
    var bodyReader = createBodyReaderForTests({});
    var documentXmlReader = new DocumentXmlReader({
        bodyReader: bodyReader
    });
    var paragraphXml = xml.element("w:p", {}, []);
    var bodyXml = xml.element("w:body2", {}, [paragraphXml]);
    var documentXml = xml.element("w:document", {}, [bodyXml]);

    assert.throws(function() {
        documentXmlReader.convertXmlToDocument(documentXml);
    }, /Could not find the body element: are you sure this is a docx file?/);
});
