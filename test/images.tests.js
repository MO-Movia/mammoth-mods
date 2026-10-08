import test from "./test.js";
import assert from "assert";
import hamjest from "hamjest";
import * as images from "../lib/images.js";
import * as documents from "../lib/documents.js";
import * as promises from "../lib/promises.js";
var assertThat = hamjest.assertThat;
var contains = hamjest.contains;
var equalTo = hamjest.equalTo;
var hasProperties = hamjest.hasProperties;

test('images.inline() should be an alias of images.imgElement()', function() {
    assert.ok(images.inline === images.imgElement);
});

test('images.dataUri() encodes images in base64', function() {
    var imageBuffer = new Buffer("abc");
    var image = new documents.Image({
        readImage: function(encoding) {
            return promises.resolve(imageBuffer.toString(encoding));
        },
        contentType: "image/jpeg"
    });

    return images.dataUri(image).then(function(result) {
        assertThat(result, contains(
            hasProperties({tag: hasProperties({attributes: {"src": "data:image/jpeg;base64,YWJj"}})})
        ));
    });
});

test('images.imgElement()', {
    'when element does not have alt text then alt attribute is not set': function() {
        var imageBuffer = new Buffer("abc");
        var image = new documents.Image({
            readImage: function(encoding) {
                return promises.resolve(imageBuffer.toString(encoding));
            },
            contentType: "image/jpeg"
        });

        var result = images.imgElement(function(image) {
            return {src: "<src>"};
        })(image);

        return result.then(function(result) {
            assertThat(result, contains(
                hasProperties({
                    tag: hasProperties({
                        attributes: equalTo({src: "<src>"})
                    })
                })
            ));
        });
    },

    'when element has alt text then alt attribute is set': function() {
        var imageBuffer = new Buffer("abc");
        var image = new documents.Image({
            readImage: function(encoding) {
                return promises.resolve(imageBuffer.toString(encoding));
            },
            contentType: "image/jpeg",
            altText: "<alt>"
        });

        var result = images.imgElement(function(image) {
            return {src: "<src>"};
        })(image);

        return result.then(function(result) {
            assertThat(result, contains(
                hasProperties({
                    tag: hasProperties({
                        attributes: equalTo({alt: "<alt>", src: "<src>"})
                    })
                })
            ));
        });
    },

    'image alt text can be overridden by alt attribute returned from function': function() {
        var imageBuffer = new Buffer("abc");
        var image = new documents.Image({
            readImage: function(encoding) {
                return promises.resolve(imageBuffer.toString(encoding));
            },
            contentType: "image/jpeg",
            altText: "<alt>"
        });

        var result = images.imgElement(function(image) {
            return {alt: "<alt override>", src: "<src>"};
        })(image);

        return result.then(function(result) {
            assertThat(result, contains(
                hasProperties({
                    tag: hasProperties({
                        attributes: equalTo({alt: "<alt override>", src: "<src>"})
                    })
                })
            ));
        });
    }
});

test("imageFilenameExtension", {
    "extension is derived from subtype of content type": function() {
        var image = new documents.Image({
            contentType: "image/gif"
        });

        var result = images.imageFilenameExtension(image);

        assertThat(result, equalTo("gif"));
    },

    "data after second slash is ignored": function() {
        var image = new documents.Image({
            contentType: "image/gif/jpeg"
        });

        var result = images.imageFilenameExtension(image);

        assertThat(result, equalTo("gif"));
    },

    "backslashes are treated as forward slashes": function() {
        var image = new documents.Image({
            contentType: "image\\gif\\..\\"
        });

        var result = images.imageFilenameExtension(image);

        assertThat(result, equalTo("gif"));
    },

    "when there is no subtype then null is returned": function() {
        var image = new documents.Image({
            contentType: "image"
        });

        var result = images.imageFilenameExtension(image);

        assertThat(result, equalTo(undefined));
    }
});
