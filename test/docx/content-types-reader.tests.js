import test from "../test.js";
import assert from "assert";
import {readContentTypesFromXml} from "../../lib/docx/content-types-reader.js";
import {Element as XmlElement} from "../../lib/xml/index.js";
test('reads default-per-extension from XML', function() {
    var contentTypes = readContentTypesFromXml(
        new XmlElement("content-types:Types", {}, [
            new XmlElement("content-types:Default", {Extension: "png", ContentType: "image/png"})
        ])
    );
    assert.equal(contentTypes.findContentType("word/media/hat.png"), "image/png");
});

test('reads overrides in preference to defaults', function() {
    var contentTypes = readContentTypesFromXml(
        new XmlElement("content-types:Types", {}, [
            new XmlElement("content-types:Default", {Extension: "png", ContentType: "image/png"}),
            new XmlElement("content-types:Override", {PartName: "/word/media/hat.png", ContentType: "image/hat"})
        ])
    );
    assert.equal(contentTypes.findContentType("word/media/hat.png"), "image/hat");
});

test('fallback content types have common image types', function() {
    var contentTypes = readContentTypesFromXml(
        new XmlElement("content-types:Types", {}, [])
    );
    assert.equal(contentTypes.findContentType("word/media/hat.png"), "image/png");
    assert.equal(contentTypes.findContentType("word/media/hat.gif"), "image/gif");
    assert.equal(contentTypes.findContentType("word/media/hat.jpg"), "image/jpeg");
    assert.equal(contentTypes.findContentType("word/media/hat.jpeg"), "image/jpeg");
    assert.equal(contentTypes.findContentType("word/media/hat.bmp"), "image/bmp");
    assert.equal(contentTypes.findContentType("word/media/hat.tif"), "image/tiff");
    assert.equal(contentTypes.findContentType("word/media/hat.tiff"), "image/tiff");
});

test('fallback content types are case insensitive on extension', function() {
    var contentTypes = readContentTypesFromXml(
        new XmlElement("content-types:Types", {}, [])
    );
    assert.equal(contentTypes.findContentType("word/media/hat.PnG"), "image/png");
});
