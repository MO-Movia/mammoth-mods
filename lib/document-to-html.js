import _ from "underscore";
import * as promises from "./promises.js";
import * as documents from "./documents.js";
import * as htmlPaths from "./styles/html-paths.js";
import * as results from "./results.js";
import * as images from "./images.js";
import * as Html from "./html/index.js";
import * as writers from "./writers/index.js";
import * as moConversions from "./mo-conversions.js";
function DocumentConverter(options) {
    return {
        convertToHtml: function(element) {
            var comments = _.indexBy(
                element.type === documents.types.document ? element.comments : [],
                "commentId"
            );
            var conversion = new DocumentConversion(options, comments);
            return conversion.convertToHtml(element);
        }
    };
}

function DocumentConversion(options, comments) {
    var noteNumber = 1;

    var noteReferences = [];

    var referencedComments = [];

    options = _.extend({ignoreEmptyParagraphs: true}, options);
    var idPrefix = options.idPrefix === undefined ? "" : options.idPrefix;
    var ignoreEmptyParagraphs = options.ignoreEmptyParagraphs;

    var defaultParagraphStyle = htmlPaths.topLevelElement("p");

    var styleMap = options.styleMap || [];

    function convertToHtml(document) {
        var messages = [];

        var html = elementToHtml(document, messages, Object.create(null));

        var deferredNodes = [];
        walkHtml(html, function(node) {
            if (node.type === "deferred") {
                deferredNodes.push(node);
            }
        });
        var deferredValues = Object.create(null);
        return promises.forEachSeries(deferredNodes, function(deferred) {
            return deferred.value().then(function(value) {
                deferredValues[deferred.id] = value;
            });
        }).then(function() {
            function replaceDeferred(nodes) {
                return flatMap(nodes, function(node) {
                    if (node.type === "deferred") {
                        return deferredValues[node.id];
                    } else if (node.children) {
                        return [
                            _.extend({}, node, {
                                children: replaceDeferred(node.children)
                            })
                        ];
                    } else {
                        return [node];
                    }
                });
            }
            var writer = writers.writer({
                prettyPrint: options.prettyPrint,
                outputFormat: options.outputFormat
            });
            Html.write(writer, Html.simplify(replaceDeferred(html)));
            return new results.Result(writer.asString(), messages);
        });
    }

    function convertElements(elements, messages, context) {
        return flatMap(elements, function(element) {
            return elementToHtml(element, messages, context);
        });
    }

    function elementToHtml(element, messages, context) {
        if (!context) {
            throw new Error("context not set");
        }
        var handler = elementConverters[element.type];
        if (handler) {
            return handler(element, messages, context);
        } else {
            return [];
        }
    }

    function convertParagraph(element, messages, context) {
        var paragraph = htmlPathForParagraph(element, messages);
        var wp = paragraph.wrap(function() {
            moConversions.moInjectNumberingSymbol(element);

            var content = convertElements(element.children, messages, context);
            if (ignoreEmptyParagraphs) {
                return content;
            } else {
                return [Html.forceWrite].concat(content);
            }
        });
        // [FS] 09-02-2023
        // Handle ListParagraph indentation, alignment and Heading style
        moConversions.moAdjustParagraphTag(element, wp);
        return moConversions.moParagraphSuffix(element, wp);
    }

    function htmlPathForParagraph(element, messages) {
        var style = findStyle(element);

        if (style) {
            return style.to;
        } else {
            if (element.styleId) {
                messages.push(unrecognisedStyleWarning("paragraph", element));
            }
            return defaultParagraphStyle;
        }
    }

    function convertRun(run, messages, context) {
        moConversions.moPropagateRunColor(run);
        var nodes = function() {
            return convertElements(run.children, messages, context);
        };
        var paths = [];
        if (run.highlight !== null) {
            var path = findHtmlPath({type: "highlight", color: run.highlight});
            if (path) {
                paths.push(path);
            }
        }
        if (run.isSmallCaps) {
            paths.push(findHtmlPathForRunProperty("smallCaps"));
        }
        if (run.isAllCaps) {
            paths.push(findHtmlPathForRunProperty("allCaps"));
        }
        if (run.isStrikethrough) {
            paths.push(findHtmlPathForRunProperty("strikethrough", "s"));
        }
        if (run.isUnderline) {
            paths.push(findHtmlPathForRunProperty("underline"));
        }
        if (run.verticalAlignment === documents.verticalAlignment.subscript) {
            paths.push(htmlPaths.element("sub", {}, {fresh: false}));
        }
        // [FS] 09-04-2024
        // Manage superscript along with infoicon.
        moConversions.moPushSuperscriptPath(run, paths);
        if (run.isItalic) {
            paths.push(findHtmlPathForRunProperty("italic", "em"));
        }
        if (run.isBold) {
            paths.push(findHtmlPathForRunProperty("bold", "strong"));
        }
        // [FS]LIC-270 Adding highlight feature to mammoth
        // Added font color feature to mammoth
        moConversions.moPushHighlightPath(run, paths);

        var stylePath = htmlPaths.empty;
        var style = findStyle(run);
        if (style) {
            stylePath = style.to;
        } else if (run.styleId) {
            messages.push(unrecognisedStyleWarning("run", run));
        }
        paths.push(stylePath);

        paths.forEach(function(path) {
            nodes = path.wrap.bind(path, nodes);
        });

        return nodes();
    }

    function findHtmlPathForRunProperty(elementType, defaultTagName) {
        var path = findHtmlPath({type: elementType});
        if (path) {
            return path;
        } else if (defaultTagName) {
            return htmlPaths.element(defaultTagName, {}, {fresh: false});
        } else {
            return htmlPaths.empty;
        }
    }

    function findHtmlPath(element, defaultPath) {
        var style = findStyle(element);
        return style ? style.to : defaultPath;
    }

    function findStyle(element) {
        for (var i = 0; i < styleMap.length; i++) {
            if (styleMap[i].from.matches(element)) {
                return styleMap[i];
            }
        }
    }

    function recoveringConvertImage(convertImage) {
        return function(image, messages) {
            return promises.try(function() {
                return convertImage(image, messages);
            }).catch(function(error) {
                messages.push(results.error(error));
                return [];
            });
        };
    }

    function noteHtmlId(note) {
        return referentHtmlId(note.noteType, note.noteId);
    }

    function noteRefHtmlId(note) {
        return referenceHtmlId(note.noteType, note.noteId);
    }

    function referentHtmlId(referenceType, referenceId) {
        return htmlId(referenceType + "-" + referenceId);
    }

    function referenceHtmlId(referenceType, referenceId) {
        return htmlId(referenceType + "-" + referenceId);
    }

    function htmlId(suffix) {
        return idPrefix + suffix;
    }

    var defaultTablePath = htmlPaths.elements([
        htmlPaths.element("table", {}, {fresh: true})
    ]);

    // [FS] 09-02-2023
    // Chapter header table row backgound and border
    function convertTable(element, messages, context) {
        var res = findHtmlPath(element, defaultTablePath).wrap(function() {
            return convertTableChildren(element, messages, context);
        });
        moConversions.moApplyTableAttributes(element, res[0].tag.attributes);
        return res;
    }

    function convertTableChildren(element, messages, context) {
        var bodyIndex = _.findIndex(element.children, function(child) {
            return !child.type === documents.types.tableRow || !child.isHeader;
        });
        if (bodyIndex === -1) {
            bodyIndex = element.children.length;
        }
        var children;
        if (bodyIndex === 0) {
            children = convertElements(
                element.children,
                messages,
                _.extend({}, context, {isTableHeader: false})
            );
        } else {
            var headRows = convertElements(
                element.children.slice(0, bodyIndex),
                messages,
                _.extend({}, context, {isTableHeader: true})
            );
            var bodyRows = convertElements(
                element.children.slice(bodyIndex),
                messages,
                _.extend({}, context, {isTableHeader: false})
            );
            children = [
                Html.freshElement("thead", {}, headRows),
                Html.freshElement("tbody", {}, bodyRows)
            ];
        }
        return [Html.forceWrite].concat(children);
    }

    function convertTableRow(element, messages, context) {
        var children = convertElements(element.children, messages, context);
        return [
            Html.freshElement("tr", {}, [Html.forceWrite].concat(children))
        ];
    }

    function convertTableCell(element, messages, context) {
        var tagName = context.isTableHeader ? "th" : "td";
        var children = convertElements(element.children, messages, context);
        var attributes = {};
        if (element.colSpan !== 1) {
            attributes.colspan = element.colSpan.toString();
        }
        if (element.rowSpan !== 1) {
            attributes.rowspan = element.rowSpan.toString();
        }
        moConversions.moApplyTableCellAttributes(element, attributes);

        return [
            Html.freshElement(
                tagName,
                attributes,
                [Html.forceWrite].concat(children)
            )
        ];
    }

    function convertCommentReference(reference, messages, context) {
        return findHtmlPath(reference, htmlPaths.ignore).wrap(function() {
            var comment = comments[reference.commentId];
            var count = referencedComments.length + 1;
            var label = "[" + commentAuthorLabel(comment) + count + "]";
            referencedComments.push({label: label, comment: comment});
            // TODO: remove duplication with note references
            return [
                Html.freshElement(
                    "a",
                    {
                        href: "#" + referentHtmlId("comment", reference.commentId),
                        id: referenceHtmlId("comment", reference.commentId)
                    },
                    [Html.text(label)]
                )
            ];
        });
    }

    function convertComment(referencedComment, messages, context) {
        // TODO: remove duplication with note references

        var label = referencedComment.label;
        var comment = referencedComment.comment;
        var body = convertElements(comment.body, messages, context).concat([
            Html.nonFreshElement("p", {}, [
                Html.text(" "),
                Html.freshElement(
                    "a",
                    {href: "#" + referenceHtmlId("comment", comment.commentId)},
                    [Html.text("↑")]
                )
            ])
        ]);

        return [
            Html.freshElement(
                "dt",
                {id: referentHtmlId("comment", comment.commentId)},
                [Html.text("Comment " + label)]
            ),
            Html.freshElement("dd", {}, body)
        ];
    }

    function convertBreak(element, messages, context) {
        return htmlPathForBreak(element).wrap(function() {
            return [];
        });
    }

    function htmlPathForBreak(element) {
        var style = findStyle(element);
        if (style) {
            return style.to;
        } else if (element.breakType === "line") {
            return htmlPaths.topLevelElement("br");
        } else {
            return htmlPaths.empty;
        }
    }

    var footNoteHeader = "footnote";
    var endNoteHeader = "endnote";

    // [FS] 09-02-2023
    // To prepare the foototes to be converted as InfoIcon
    /**
   *
   * @param Header: Note Type
   * @param noteReferences: Array of Notes
   * @param document:The current document
   * @param messages:Messages to be shown at bottom
   * @param options
   */
    function prepareNote(Header, noteReferences, document, messages, options) {
        return moConversions.moPrepareNote(
            Header,
            noteReferences,
            document,
            messages,
            options,
            convertElements
        );
    }

    var elementConverters = {
        "document": function(document, messages, context) {
            var children = convertElements(document.children, messages, context);
            var footnotesNodes = prepareNote(
                footNoteHeader,
                noteReferences,
                document,
                messages,
                context
            );
            var endnotesNodes = prepareNote(
                endNoteHeader,
                noteReferences,
                document,
                messages,
                context
            );
            // [FS] 09-02-2023
            // To prepare the foototes to be converted as InfoIcon
            return children.concat(
                moConversions.moNoteContainers(footnotesNodes, endnotesNodes).concat([
                    Html.freshElement(
                        "dl",
                        {},
                        flatMap(referencedComments, function(referencedComment) {
                            return convertComment(referencedComment, messages, context);
                        })
                    )
                ])
            );
        },
        "paragraph": convertParagraph,
        "run": convertRun,
        "text": function(element, messages, context) {
            return [Html.text(element.value)];
        },
        "tab": function(element, messages, context) {
            return [Html.text("\t")];
        },
        "hyperlink": function(element, messages, context) {
            var href = element.anchor ? "#" + htmlId(element.anchor) : element.href;
            var attributes = moConversions.moHyperlinkAttributes(element, href);
            if (element.targetFrame != null) {
                attributes.target = element.targetFrame;
            }

            var children = convertElements(element.children, messages, context);
            return [Html.nonFreshElement("a", attributes, children)];
        },
        "checkbox": function(element) {
            var attributes = {type: "checkbox"};
            if (element.checked) {
                attributes["checked"] = "checked";
            }
            return [Html.freshElement("input", attributes)];
        },
        "bookmarkStart": function(element, messages, context) {
            var anchor = Html.freshElement("a", {
                id: htmlId(element.name)
            }, [Html.forceWrite]);
            return [anchor];
        },
        "noteReference": function(element, messages, context) {
            noteReferences.push(element);
            var anchor = Html.freshElement(
                "span",
                {
                    id: noteRefHtmlId(element)
                },
                [Html.text("[" + noteNumber++ + "]")]
            );
            // [FS] 09-02-2023
            // To prepare the foototes to be converted as InfoIcon
            return [
                Html.freshElement("sup", moConversions.moNoteReferenceAttributes(element), [
                    anchor
                ])
            ];
        },
        "note": function(element, messages, context) {
            // [FS] 09-02-2023
            // To prepare the foototes to be converted as InfoIcon
            moConversions.moMarkInfoIconNote(element);
            var children = convertElements(element.body, messages, context);
            var backLink = Html.elementWithTag(
                htmlPaths.element("p", {}, {fresh: false}),
                [Html.text(" ")]
            );
            var body = children.concat([backLink]);

            return Html.freshElement("li", {id: noteHtmlId(element)}, body);
        },
        "commentReference": convertCommentReference,
        "comment": convertComment,
        "image": deferredConversion(
            recoveringConvertImage(options.convertImage || images.dataUri)
        ),
        "table": convertTable,
        "tableRow": convertTableRow,
        "tableCell": convertTableCell,
        "break": convertBreak
    };
    return {
        convertToHtml: convertToHtml
    };
}

var deferredId = 1;

function deferredConversion(func) {
    return function(element, messages, context) {
        return [
            {
                type: "deferred",
                id: deferredId++,
                value: function() {
                    return func(element, messages, context);
                }
            }
        ];
    };
}

function unrecognisedStyleWarning(type, element) {
    return results.warning(
        "Unrecognised " +
    type +
    " style: '" +
    element.styleName +
    "'" +
    " (Style ID: " +
    element.styleId +
    ")"
    );
}

function flatMap(values, func) {
    return _.flatten(values.map(func), true);
}

function walkHtml(nodes, callback) {
    nodes.forEach(function(node) {
        callback(node);
        if (node.children) {
            walkHtml(node.children, callback);
        }
    });
}

var commentAuthorLabel = (function  commentAuthorLabel(comment) {
    return comment.authorInitials || "";
});

export {commentAuthorLabel, DocumentConverter};
