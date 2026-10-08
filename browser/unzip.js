import * as promises from "../lib/promises.js";
import * as zipfile from "../lib/zipfile.js";
function openZip(options) {
    if (options.arrayBuffer) {
        return promises.resolve(zipfile.openArrayBuffer(options.arrayBuffer));
    } else {
        return promises.reject(new Error("Could not find file in options"));
    }
}

export {openZip};
