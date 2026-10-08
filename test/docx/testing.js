import {createBodyReader} from "../../lib/docx/body-reader.js";
import {defaultNumbering} from "../../lib/docx/numbering-xml.js";
import {Styles} from "../../lib/docx/styles-reader.js";
function createBodyReaderForTests(options) {
    options = Object.create(options || {});
    options.styles = options.styles || new Styles({}, {});
    options.numbering = options.numbering || defaultNumbering;
    return createBodyReader(options);
}

export {createBodyReaderForTests};
