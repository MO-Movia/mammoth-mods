import path from "path";
import _ from "underscore";
import * as fs from "../lib/fs.js";
import * as promises from "../lib/promises.js";
import {fileURLToPath} from "url";

var __filename = fileURLToPath(import.meta.url);
var __dirname = path.dirname(__filename);
function testPath(filename) {
    return path.join(__dirname, "test-data", filename);
}

function testData(testDataPath) {
    var fullPath = testPath(testDataPath);
    return fs.readFile(fullPath, "utf-8");
}

function createFakeDocxFile(files) {
    function exists(path) {
        return !!files[path];
    }

    return {
        read: createRead(files),
        exists: exists
    };
}

function createFakeFiles(files) {
    return {
        read: createRead(files)
    };
}

function createRead(files) {
    function read(path, encoding) {
        return promises.resolve(files[path], function(buffer) {
            if (_.isString(buffer)) {
                buffer = new Buffer(buffer);
            }

            if (!Buffer.isBuffer(buffer)) {
                return promises.reject(new Error("file was not a buffer"));
            } else if (encoding) {
                return promises.resolve(buffer.toString(encoding));
            } else {
                return promises.resolve(buffer.buffer);
            }
        });
    }
    return read;
}

export {testPath, testData, createFakeDocxFile, createFakeFiles};
