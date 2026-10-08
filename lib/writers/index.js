var htmlWriter = require("./html-writer");

exports.writer = writer;


function writer(options) {
    return htmlWriter.writer(options);
}
