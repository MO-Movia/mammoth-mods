import fs from "fs";
import * as promises from "./promises.js";
function readFile(path, options) {
    return promises.resolve(fs.promises.readFile(path, options));
}

function writeFile(path, data, options) {
    return promises.resolve(fs.promises.writeFile(path, data, options));
}

var createWriteStream = fs.createWriteStream.bind(fs);

var readFileSync = fs.readFileSync.bind(fs);

export {readFile, writeFile, createWriteStream, readFileSync};
