import {warning} from "../results.js";

// MO-Movia fork extensions: custom paragraph/run/style property readers.
// Extracted from body-reader.js so that upstream merges only touch the
// call sites there.


// [FS]LIC-270 Adding highlight feature to mammoth; also run text color.
function readMoRunProperties(element) {
    return {
        color: element.firstOrEmpty("w:color").attributes["w:val"],
        isHighlight: readBooleanElement(element.first("w:highlight")),
        highlightColor: element.firstOrEmpty("w:highlight").attributes["w:val"]
    };
}

function readBooleanElement(element) {
    if (element) {
        var value = element.attributes["w:val"];
        return value !== "false" && value !== "0";
    } else {
        return false;
    }
}

// [FS] 09-01-2023
// Adding Horizontal rule for Bottom Border
function readMoBottomBorderAttributes(element) {
    if (element.children) {
        for (var i = 0; i < element.children.length; i++) {
            if (element.children[i].name === "w:pBdr") {
                for (var j = 0; j < element.children[i].children.length; j++) {
                    if (element.children[i].children[j].name === "w:bottom") {
                        return element.children[i].children[j].attributes;
                    } else {
                        return null;
                    }
                }
            }
        }
    }
}

// [FS]
// ISSUE FIX: Font applied for all paragraphs
function resetMoImageRunProperties(properties, children) {
    if (
        children.length <= 1 &&
        children[0] &&
        children[0].type === "image" &&
        properties
    ) {
        properties.font = null;
        properties.fontSize = null;
        properties.isBold = false;
        properties.isUnderline = false;
        properties.isItalic = false;
        properties.isStrikethrough = false;
        properties.isAllCaps = false;
        properties.isSmallCaps = false;
        properties.color = null;
        properties.isHighlight = null;
    }
}

// [FS] 09-02-2023
// Set paragraphs attribute to its textbox attributes to
function setMoTextBoxAttributes(element) {
    if (element.children) {
        for (var index = 0; index < element.children.length; index++) {
            if (element.children[index].name === "v:textbox") {
                element.children[index].attributes.textbox = element.attributes;
                return element.attributes;
            } else {
                var attr = setMoTextBoxAttributes(element.children[index]);
                if (attr) {
                    return attr;
                }
            }
        }
    }
}

// [FS] 11-03-2023
// Custom style resolution: emits "parsedStyles" warnings for matched and
// unmatched styles.
// Returns {styleId, name, messages}; the caller wraps it in a ReadResult.
function readMoStyleInfo(element, styleTagName, styleType, findStyleById) {
    var messages = [];
    var styleElement = element.first(styleTagName);
    var styleId = null;
    var name = null;
    if (styleElement) {
        styleId = styleElement.attributes["w:val"];

        if (styleId) {
            var style = findStyleById(styleId);
            if (style) {
                name = style.name;
                messages.push(parsedStyles(style));
            } else {
                messages.push(parsedStyles(styleId));
                messages.push(undefinedStyleWarning(styleType, styleId));
            }
        }
    }
    return {styleId: styleId, name: name, messages: messages};
}

function undefinedStyleWarning(type, styleId) {
    return warning(
        type +
            " style with ID " +
            styleId +
            " was referenced but not defined in the document"
    );
}

function parsedStyles(elementStyle) {
    var style = "";
    if (elementStyle.name) {
        style = elementStyle.name;
    } else if (elementStyle.styleId) {
        style = elementStyle.styleId;
    }
    return warning(
        " parsedStyles: '" +
            elementStyle.name +
            "'" +
            " (Style ID: " +
            style +
            ")"
    );
}

export {readMoRunProperties, readMoBottomBorderAttributes, resetMoImageRunProperties, setMoTextBoxAttributes, readMoStyleInfo};
