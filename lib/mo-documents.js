
// MO-Movia fork extensions: document-model property helpers.
// Extracted from documents.js so that upstream merges only touch the
// call sites there.

// [FS] 16-01-2023
// To check whether the paragraph has Bottom Border
function moIsBottomBorder(properties) {
    return (
        properties.type === "paragraphProperties" &&
        properties.bottomBorderAttributes !== undefined
    );
}

// Word uses "auto" for the automatic text color; emit black instead.
function moAutoColor(properties) {
    if (properties.color === "auto") {
        return "#000000";
    } else {
        return "#" + properties.color;
    }
}

export {moIsBottomBorder, moAutoColor};
