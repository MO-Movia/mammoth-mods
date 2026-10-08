import xmldom from "@xmldom/xmldom";
import dom from "@xmldom/xmldom/lib/dom.js";
function parseFromString(string) {
    var error = null;

    var domParser = new xmldom.DOMParser({
        errorHandler: function(level, message) {
            error = {level: level, message: message};
        }
    });

    var document = domParser.parseFromString(string);

    if (error === null) {
        return document;
    } else {
        throw new Error(error.level + ": " + error.message);
    }
}

var Node = dom.Node;

export {parseFromString, Node};
