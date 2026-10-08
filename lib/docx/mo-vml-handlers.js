// MO-Movia fork extensions: VML element readers (vignettes, shapes, textboxes).
// Extracted from body-reader.js so that upstream merges only touch the
// _.extend() call site there.

exports.createMoVmlHandlers = createMoVmlHandlers;

function createMoVmlHandlers(deps) {
    var readChildElements = deps.readChildElements;
    var elementResult = deps.elementResult;

    return {
        "v:roundrect": readRoundrectVignette,
        "v:shape": readShape,
        "v:textbox": readTextbox
    };

    function readRoundrectVignette(element) {
        // Extract styling from v:roundrect attributes
        var rawFill = element.attributes.fillcolor || "#DCE6F2";
        var bracketIndex = rawFill.indexOf('[');
        var fillcolor = bracketIndex !== -1
            ? rawFill.substring(0, bracketIndex).trim()
            : rawFill.trim();
        var arcsize = element.attributes.arcsize || "0.0967592592592593";

        // Extract v:stroke properties
        var strokeElement = element.first("v:stroke");
        var strokeWeight = strokeElement ? (strokeElement.attributes.weight || "1pt") : "1pt";
        var strokeColor = strokeElement ? (strokeElement.attributes.color || "#17375E") : "#17375E";

        // Extract v:fill properties
        var fillElement = element.first("v:fill");
        var fillOn = fillElement ? (fillElement.attributes.on !== "f") : true;

        // Extract dimensions from style attribute
        var style = element.attributes.style || "";
        var width = "";
        var height = "";

        if (style) {
            var styleVals = style.split(";");
            styleVals.forEach(function(val) {
                if (val.startsWith("width")) {
                    width = val.split(":")[1].trim();
                }
                if (val.startsWith("height")) {
                    height = val.split(":")[1].trim();
                }
            });
        }

        // Build the vignette style string
        var vignetteStyle = "";

        // Background color (only if fill is on)
        if (fillOn && fillcolor && fillcolor !== "none") {
            vignetteStyle += "background-color:" + fillcolor + ";";
        }

        // Border
        if (strokeWeight && strokeWeight !== "0pt") {
            vignetteStyle += "border:" + strokeWeight + " solid " + strokeColor + ";";
        }

        // Calculate border-radius from arcsize
        var borderRadiusValue = "15px";
        if (arcsize) {
            var arcsizeNum = parseFloat(arcsize);
            if (arcsizeNum > 0) {
                borderRadiusValue = Math.round(arcsizeNum * 200) + "px";
            }
        }
        vignetteStyle += "border-radius:" + borderRadiusValue + ";";

        // Dimensions
        if (width) {
            vignetteStyle += "width:" + width + ";";
        }
        if (height) {
            vignetteStyle += "min-height:" + height + ";";
        }

        // Box model and display
        vignetteStyle += "padding:15px;";
        vignetteStyle += "margin:10px 0;";
        vignetteStyle += "display:block;";
        vignetteStyle += "box-sizing:border-box;";

        // Read the children (content inside the vignette)
        var children = readChildElements(element);

        // Create a wrapper div with vignette styling
        var wrapper = {
            type: "paragraph",
            children: children.value || [],
            styleId: null,
            styleName: "vignette",
            numbering: null,
            alignment: null,
            indent: {
                start: null,
                end: null,
                firstLine: null,
                hanging: null
            },
            attributes: {
                style: vignetteStyle,
                class: "vignette-box"
            },
            attrs: null
        };

        // Return the wrapped result
        return elementResult(wrapper);
    }

    // [FS] 09-02-2023
    // Get the style properties of shape element
    function readShape(element) {
        var width = "";
        var height = "";
        var style = "";
        if (element.attributes.id && element.attributes.id.includes("Pic")) {
            if (element.attributes.style) {
                var styleVals = element.attributes.style.split(";");
                styleVals.forEach(function(val) {
                    if (val.startsWith("width")) {
                        var w = val.split(":")[1];
                        width = "width:" + twipToPoint(w) + ";";
                        style = style + width;
                    }
                    if (val.startsWith("height")) {
                        var h = val.split(":")[1];
                        height = "height:" + twipToPoint(h) + ";";
                        style = style + height;
                    }
                });
                for (var i = 0; i < element.children.length; i++) {
                    if (element.children[i].name === "v:imagedata") {
                        element.children[i].attributes.style = width + height;
                    }
                }
            }
        }
        var shape = readChildElements(element);
        if (element.attributes.id && element.attributes.id.includes("Pic")) {
            var np = {
                alignment: null,
                children: [],
                indent: {},
                numbering: null,
                styleId: null,
                styleName: null,
                attributes: {},
                type: "paragraph"
            };
            np.children = shape.value;
            np.attributes.style = style;
            shape.value = np;
            shape._result.value.element = np;
        }

        return shape;
    }

    // [FS] 09-02-2023
    // Manipulate the textbox properties
    function readTextbox(element) {
        var txtBox = readChildElements(element);
        var style = "";
        var width = "";
        var height = "";
        if (element.attributes.textbox) {
            var styleVals = element.attributes.textbox.style.split(";");
            styleVals.forEach(function(val) {
                if (val.startsWith("width")) {
                    var w = val.split(":")[1];
                    width = "width:" + twipToPoint(w) + ";";
                }
                if (val.startsWith("height")) {
                    var h = val.split(":")[1];
                    height = "height:" + twipToPoint(h) + ";";
                }
            });

            var txtBoxStyle = "";
            var np = {
                alignment: null,
                children: [],
                indent: {},
                numbering: null,
                styleId: null,
                styleName: null,
                attributes: {},
                type: "paragraph"
            };
            np.children = txtBox.value;
            txtBoxStyle = element.attributes.textbox.fillcolor
                ? "background-color:" +
                  element.attributes.textbox.fillcolor.split(" ")[0] +
                  ";"
                : "";
            if (element.attributes.textbox.strokeweight) {
                if (!element.attributes.textbox.strokecolor) {
                    element.attributes.textbox.strokecolor = "black";
                }
                txtBoxStyle =
                    txtBoxStyle +
                    "border:" +
                    element.attributes.textbox.strokeweight +
                    " solid " +
                    element.attributes.textbox.strokecolor.split(" ")[0] +
                    ";";
            }

            if (txtBoxStyle != "") {
                np.attributes.style = "display:inline-block;" + txtBoxStyle + style;
            } else {
                np.attributes.style =
                    "border: 1px solid black;display:inline-block;" + style;
            }

            txtBox.value = np;
            txtBox._result.value.element = np;
        }

        return txtBox;
    }

    function twipToPoint(val) {
        if (!isNaN(val)) {
            return val / 20 + "pt";
        } else {
            if (val && val.includes("pt")) {
                return val;
            }
        }
    }
}
