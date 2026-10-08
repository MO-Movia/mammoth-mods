import _ from "underscore";
import * as docxReader from "./docx/docx-reader.js";
import * as docxStyleMap from "./docx/style-map.js";
import * as promises from "./promises.js";
import * as unzip from "./unzip.js";
import {DocumentConverter} from "./document-to-html.js";
import {readStyle} from "./style-reader.js";
import {readOptions} from "./options-reader.js";
import {Result} from "./results.js";


function convertToHtml(input, options) {
    return convert(input, options);
}

function convert(input, options) {
    options = readOptions(options);

    var result = unzip.openZip(input)
        .then(function(docxFile) {
            return docxStyleMap.readStyleMap(docxFile).then(function(styleMap) {
                options.embeddedStyleMap = styleMap;
                return docxFile;
            });
        })
        .then(function(docxFile) {
            return docxReader.read(docxFile, input, options)
                .then(function(documentResult) {
                    return documentResult.map(options.transformDocument);
                })
                .then(function(documentResult) {
                    return convertDocumentToHtml(documentResult, options);
                });
        });

    return promises.toExternalPromise(result);
}

function convertDocumentToHtml(documentResult, options) {
    var styleMapResult = parseStyleMap(options.readStyleMap());
    var parsedOptions = _.extend({}, options, {
        styleMap: styleMapResult.value
    });
    var documentConverter = new DocumentConverter(parsedOptions);

    return documentResult.flatMapThen(function(document) {
        return styleMapResult.flatMapThen(function(styleMap) {
            return documentConverter.convertToHtml(document);
        });
    });
}

function parseStyleMap(styleMap) {
    return Result.combine((styleMap || []).map(readStyle))
        .map(function(styleMap) {
            return styleMap.filter(function(styleMapping) {
                return !!styleMapping;
            });
        });
}

export {convertToHtml, convert};

export default {convertToHtml: convertToHtml, convert: convert};
