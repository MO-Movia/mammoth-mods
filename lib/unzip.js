import * as fs from "./fs.js";
import * as promises from "./promises.js";
import * as zipfile from "./zipfile.js";
function openZip(options) {
    if (options.path) {
        return fs.readFile(options.path).then(zipfile.openArrayBuffer);
    } else if (options.buffer) {
        return promises.resolve(zipfile.openArrayBuffer(options.buffer));
    } else if (options.file) {
        return promises.resolve(options.file);
    } else {
        return promises.reject(new Error("Could not find file in options"));
    }
}

export {openZip};
