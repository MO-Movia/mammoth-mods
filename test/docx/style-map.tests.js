var assert = require("assert");

var JSZip = require("jszip");

var zipfile = require("../../lib/zipfile");
var styleMap = require("../../lib/docx/style-map");
var test = require("../test")(module);

test('reading embedded style map on document without embedded style map returns null', function() {
    return normalDocx().then(function(zip) {
        return styleMap.readStyleMap(zip).then(function(contents) {
            assert.equal(contents, null);
        });
    });
});

function normalDocx() {
    var zip = new JSZip();
    var originalRelationshipsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
        '</Relationships>';
    var originalContentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="png" ContentType="image/png"/>' +
        '</Types>';
    zip.file("word/_rels/document.xml.rels", originalRelationshipsXml);
    zip.file("[Content_Types].xml", originalContentTypesXml);
    return zip.generateAsync({type: "arraybuffer"}).then(function(buffer) {
        return zipfile.openArrayBuffer(buffer);
    });
}
