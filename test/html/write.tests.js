import test from "../test.js";
import assert from "assert";
import * as html from "../../lib/html/index.js";
import * as writers from "../../lib/writers/index.js";
test("text is HTML escaped", function() {
    assert.equal(
        generateString(html.text("<>&")),
        "&lt;&gt;&amp;");
});

test("double quotes outside of attributes are not escaped", function() {
    assert.equal(
        generateString(html.text('"')),
        '"');
});

test("element attributes are HTML escaped", function() {
    assert.equal(
        generateString(html.freshElement("p", {"x": "<"})),
        '<p x="&lt;"></p>');
});

test("double quotes inside attributes are escaped", function() {
    assert.equal(
        generateString(html.freshElement("p", {"x": '"'})),
        '<p x="&quot;"></p>');
});

test("element children are written", function() {
    assert.equal(
        generateString(html.freshElement("p", {}, [html.text("Hello")])),
        '<p>Hello</p>');
});

function generateString(node) {
    var writer = writers.writer();
    html.write(writer, [node]);
    return writer.asString();
}
