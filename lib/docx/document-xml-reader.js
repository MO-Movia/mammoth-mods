import {Result} from "../results.js";
import * as documents from "../documents.js";
import * as moChapterHeader from "./mo-chapter-header.js";


// [FS] 16-01-2023
// Chapter Header Images inside Table


function DocumentXmlReader(options) {
    var bodyReader = options.bodyReader;

    function convertXmlToDocument(element) {
    // [FS] 16-01-2023
    // Chapter Header Images inside Table
        var body = moChapterHeader.checkGroupStyleName(element.first('w:body'));
        if (body == null) {
            throw new Error(
                'Could not find the body element: are you sure this is a docx file?'
            );
        }

        var result = bodyReader
            .readXmlElements(body.children)
            .map(function(children) {
                return new documents.Document(children, {
                    notes: options.notes,
                    comments: options.comments
                });
            });
        return new Result(result.value, result.messages);
    }

    return {
        convertXmlToDocument: convertXmlToDocument
    };
}

export {DocumentXmlReader};
