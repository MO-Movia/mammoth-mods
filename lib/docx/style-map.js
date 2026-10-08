var promises = require("../promises");

exports.readStyleMap = readStyleMap;

var styleMapPath = "mammoth/style-map";

function readStyleMap(docxFile) {
    if (docxFile.exists(styleMapPath)) {
        return docxFile.read(styleMapPath, "utf8");
    } else {
        return promises.resolve(null);
    }
}
