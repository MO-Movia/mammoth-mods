import * as ast from "./ast.js";
export {default as simplify} from "./simplify.js";
var freshElement = ast.freshElement;
var nonFreshElement = ast.nonFreshElement;
var elementWithTag = ast.elementWithTag;
var text = ast.text;
var forceWrite = ast.forceWrite;

function write(writer, nodes) {
    nodes.forEach(function(node) {
        writeNode(writer, node);
    });
}

function writeNode(writer, node) {
    toStrings[node.type](writer, node);
}

var toStrings = {
    element: generateElementString,
    text: generateTextString,
    forceWrite: function() { }
};

function generateElementString(writer, node) {
    if (ast.isVoidElement(node)) {
        writer.selfClosing(node.tag.tagName, node.tag.attributes);
    } else {
        writer.open(node.tag.tagName, node.tag.attributes);
        write(writer, node.children);
        writer.close(node.tag.tagName);
    }
}

function generateTextString(writer, node) {
    writer.text(node.value);
}

export {write, freshElement, nonFreshElement, elementWithTag, text, forceWrite};
