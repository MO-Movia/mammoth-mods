// MO-Movia fork extensions: document-model -> HTML conversion additions.
// Extracted from document-to-html.js so that upstream merges only touch the
// call sites there.

var documents = require("./documents");
var Html = require("./html");
var htmlPaths = require("./styles/html-paths");

exports.moInjectNumberingSymbol = moInjectNumberingSymbol;
exports.moAdjustParagraphTag = moAdjustParagraphTag;
exports.moParagraphSuffix = moParagraphSuffix;
exports.moPropagateRunColor = moPropagateRunColor;
exports.moPushSuperscriptPath = moPushSuperscriptPath;
exports.moPushHighlightPath = moPushHighlightPath;
exports.moApplyTableAttributes = moApplyTableAttributes;
exports.moApplyTableCellAttributes = moApplyTableCellAttributes;
exports.moPrepareNote = moPrepareNote;
exports.moNoteContainers = moNoteContainers;
exports.moNoteReferenceAttributes = moNoteReferenceAttributes;
exports.moMarkInfoIconNote = moMarkInfoIconNote;
exports.moHyperlinkAttributes = moHyperlinkAttributes;

// For unordered lists, inject the numbering symbol (bullet character) as a
// run at the start of the paragraph so it survives into the HTML output.
function moInjectNumberingSymbol(element) {
    if (element.numbering && !element.numbering.isOrdered) {
        var newChildren = {
            type: "run",
            children: [{type: "text", value: element.numbering.symbol}],
            font: element.numbering.font,
            color: element.numbering.color
        };
        element.children.unshift(newChildren);
    }
}

// [FS] 09-02-2023
// Handle ListParagraph indentation, alignment and Heading style, plus
// vignette/group paragraph attribute wrappers.
function moAdjustParagraphTag(element, wp) {
    if (
        element.styleId == "ListParagraph" &&
        element.numbering &&
        element.numbering.level
    ) {
        var tagAttributes = {
            class: wp[0].tag.attributes.class,
            "data-indent": parseInt(element.numbering.level, 10).toString(),
            "list-style-level": (parseInt(element.numbering.level, 10) + 1).toString()
        };

        wp[0].tag = htmlPaths.element("p", tagAttributes, {
            fresh: true
        });
    }

    if (element.alignment || element.styleId) {
        var attr = {};
        if (element.alignment) {
            attr.align = element.alignment;
        }
        if (element.styleId === "Chapter Header") {
            element.styleName = element.styleId;
        }
        if (element.styleName) {
            if (element.styleName == "heading 1") {
                attr.className = "Chapter Header";
            }
            if (element.styleId === 'Heading2') {
                attr.className = 'Header 2';
            } else if (element.styleId == "ListParagraph") {
                attr.className = "List-style";
            } else {
                attr.className = element.styleName;
            }
        }
        if (wp[0] !== undefined) {
            var attributes = Object.assign({}, wp[0].tag.attributes, attr);
            var tag = htmlPaths.element("p", attributes, {fresh: true});
            wp[0].tag = tag;
        }
    }

    if (element.attributes && element.attributes.style) {
        tag = htmlPaths.element("span", element.attributes, {fresh: true});
        wp[0].tag = tag;
    }

    if (element.attributes && element.attributes.name === "group") {
        tag = htmlPaths.element("div", element.attributes, {fresh: true});
        wp[0].tag = tag;
    }
}

// [FS] 09-01-2023
// Paragraphs with a bottom border also emit an <hr> carrying the border color.
function moParagraphSuffix(element, wp) {
    if (element.attrs && element.children.length > 0) {
        var hrColor =
            element.attrs["w:color"] === "auto"
                ? "#000000"
                : element.attrs["w:color"];
        return wp.concat(Html.freshElement("hr", {color: hrColor}, []));
    } else {
        return wp;
    }
}

// Propagate run color onto hyperlink children.
function moPropagateRunColor(run) {
    for (var i = 0; i < run.children.length; i++) {
        if (run.children[i].type === "hyperlink" && run.color) {
            run.children[i]["color"] = run.color;
        }
    }
}

// [FS] 09-04-2024
// Manage superscript along with infoicon: only emit <sup> when the run's
// first child is text.
function moPushSuperscriptPath(run, paths) {
    if (run.verticalAlignment === documents.verticalAlignment.superscript) {
        if (run.children.length > 0 && run.children[0].type === "text") {
            paths.push(htmlPaths.element("sup", {}, {fresh: false}));
        }
    }
}

// [FS]LIC-270 Adding highlight feature to mammoth
// Added font color feature to mammoth
function moPushHighlightPath(run, paths) {
    if (run.isHighlight || run.color) {
        paths.push(htmlPaths.element("mark-text-highlight", {'highlight-color': run.highlightColor, 'color': run.color ? run.color : '#000'}, {fresh: false}));
    }
}

// [FS] 09-02-2023
// Chapter header table row background and border.
function moApplyTableAttributes(element, tagAttributes) {
    if (element.isCusTable) {
        tagAttributes.style =
            "border-collapse:collapse;background-color:#d8d8d8;";
    } else {
        tagAttributes.style = "border-collapse:collapse";
    }
    tagAttributes["border"] = "1px solid #000000";
}

function moApplyTableCellAttributes(element, attributes) {
    if (element.isLogoImg) {
        attributes.style = "mix-blend-mode: multiply;";
    }
}

// [FS] 09-02-2023
// To prepare the footnotes to be converted as InfoIcon.
// Returns converted note elements for the given note type.
function moPrepareNote(Header, noteReferences, document, messages, options, convertElements) {
    var filter = [];
    for (var i = 0; i < noteReferences.length; i++) {
        if (noteReferences[i].noteType === Header) {
            filter.push(noteReferences[i]);
        }
    }
    var NoteData = filter.map(function(noteReference) {
        return document.notes.resolve(noteReference);
    });
    return convertElements(NoteData, messages, options);
}

// [FS] 09-02-2023
// Footnotes and endnotes are emitted as <ol> containers with the ids the
// downstream converter expects (footnotes become InfoIcons).
function moNoteContainers(footnotesNodes, endnotesNodes) {
    return [
        Html.freshElement("ol", {id: "infoIcon"}, footnotesNodes),
        Html.freshElement("ol", {id: "endNotes"}, endnotesNodes)
    ];
}

// [FS] 09-02-2023
// Footnote references are wrapped in <sup id="infoIcon">.
function moNoteReferenceAttributes(element) {
    if (element.noteType && element.noteType == "footnote") {
        return {id: "infoIcon"};
    } else {
        return {};
    }
}

// [FS] 09-02-2023
// Footnote bodies are flagged so they are converted as InfoIcons.
function moMarkInfoIconNote(element) {
    if (element.noteType && element.noteType == "footnote") {
        element.body[0]["isInfoIcon"] = true;
    }
}

// Hyperlinks carry the propagated run color as an attribute.
function moHyperlinkAttributes(element, href) {
    if (element.color) {
        return {href: href, color: element.color};
    } else {
        return {href: href};
    }
}
