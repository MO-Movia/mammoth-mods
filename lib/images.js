import _ from "underscore";
import * as promises from "./promises.js";
import * as Html from "./html/index.js";
function imgElement(func) {
    return function(element, messages) {
        return promises.resolve(func(element)).then(function(result) {
            var attributes = {};
            if (element.altText) {
                attributes.alt = element.altText;
            }
            _.extend(attributes, result);

            return [Html.freshElement("img", attributes)];
        });
    };
}

// Undocumented, but retained for backwards-compatibility with 0.3.x
var inline = imgElement;

var dataUri = imgElement(function(element) {
    return element.read("base64").then(function(imageBuffer) {
        // [FS] 25-01-2023
        // To check whether the paragraph has Bottom Border
        var imgSrc = "data:" + element.contentType + ";base64," + imageBuffer;
        var ImgID = element.imageId;
        var retVal = element.imageId ? {src: imgSrc, id: ImgID} : {src: imgSrc};
        return retVal;

    });
});

function imageFilenameExtension(image) {
    return image.contentType.split(/\/|\\/)[1];
}

export {imgElement, imageFilenameExtension, inline, dataUri};
