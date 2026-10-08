import test from "../test.js";
import assert from "assert";
import * as xml from "../../lib/xml/index.js";
import * as officeXmlReader from "../../lib/docx/office-xml-reader.js";
test("mc:AlternateContent", {
    "when mc:Fallback is present then mc:Fallback is read": function() {
        var xmlString =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<numbering xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
            '<mc:AlternateContent>' +
            '<mc:Choice Requires="w14">' +
            '<choice/>' +
            '</mc:Choice>' +
            '<mc:Fallback>' +
            '<fallback/>' +
            '</mc:Fallback>' +
            '</mc:AlternateContent>' +
            '</numbering>';
        return officeXmlReader.read(xmlString).then(function(element) {
            assert.deepEqual(element.children, [xml.element("fallback")]);
        });
    },

    "when mc:Fallback is not present then element is ignored": function() {
        var xmlString =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<numbering xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
            '<mc:AlternateContent>' +
            '<mc:Choice Requires="w14">' +
            '<choice/>' +
            '</mc:Choice>' +
            '</mc:AlternateContent>' +
            '</numbering>';
        return officeXmlReader.read(xmlString).then(function(element) {
            assert.deepEqual(element.children, []);
        });
    }
});
