import * as documents from "../documents.js";
import {Result} from "../results.js";


function createCommentsReader(bodyReader) {
    function readCommentsXml(element) {
        return Result.combine(element.getElementsByTagName("w:comment")
            .map(readCommentElement));
    }

    function readCommentElement(element) {
        var id = element.attributes["w:id"];

        function readOptionalAttribute(name) {
            return (element.attributes[name] || "").trim() || null;
        }

        return bodyReader.readXmlElements(element.children)
            .map(function(body) {
                return documents.comment({
                    commentId: id,
                    body: body,
                    authorName: readOptionalAttribute("w:author"),
                    authorInitials: readOptionalAttribute("w:initials")
                });
            });
    }
    
    return readCommentsXml;
}

export {createCommentsReader};
