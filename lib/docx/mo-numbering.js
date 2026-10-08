// MO-Movia fork extensions: extra numbering level properties (bullet symbol,
// font and color) used by downstream bullet handling.
// Extracted from numbering-xml.js so upstream merges only touch the call site.

exports.readMoNumberingLevelExtras = readMoNumberingLevelExtras;

function readMoNumberingLevelExtras(levelElement) {
    var font = '';
    var color = '';
    var symbol = levelElement.firstOrEmpty("w:lvlText").attributes["w:val"];
    if (levelElement.first("w:rPr")) {
        font = levelElement.first("w:rPr").firstOrEmpty("w:rFonts").attributes["w:ascii"];
        color = levelElement.first("w:rPr").firstOrEmpty("w:color").attributes["w:val"];
    }
    return {
        font: font,
        symbol: symbol,
        color: color ? ("#" + color) : ''
    };
}
