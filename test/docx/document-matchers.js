import hamjest from "hamjest";
import _ from "underscore";
import * as documents from "../../lib/documents.js";
var isEmptyRun = isRun({children: []});

function isRun(properties) {
    return isDocumentElement(documents.types.run, properties);
}

function isText(text) {
    return isDocumentElement(documents.types.text, {value: text});
}

function isCheckbox(properties) {
    return isDocumentElement(documents.types.checkbox, properties);
}

function isHyperlink(properties) {
    return isDocumentElement(documents.types.hyperlink, properties);
}

function isTable(options) {
    return isDocumentElement(documents.types.table, options);
}

function isRow(options) {
    return isDocumentElement(documents.types.tableRow, options);
}

function isDocumentElement(type, properties) {
    return hamjest.hasProperties(_.extend({type: hamjest.equalTo(type)}, properties));
}

export {isRun, isText, isCheckbox, isHyperlink, isTable, isRow, isEmptyRun};
