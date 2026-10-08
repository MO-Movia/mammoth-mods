import test from "./test.js";
import assert from "assert";
import path from "path";
import * as fs from "../lib/fs.js";
import * as unzip from "../lib/unzip.js";
import {fileURLToPath} from "url";

var __filename = fileURLToPath(import.meta.url);
var __dirname = path.dirname(__filename);
test("unzip fails if given empty object", function() {
    return unzip.openZip({}).then(function() {
        assert.ok(false, "Expected failure");
    }, function(error) {
        assert.equal("Could not find file in options", error.message);
    });
});

test("unzip can open local zip file", function() {
    var zipPath = path.join(__dirname, "test-data/hello.zip");
    return unzip.openZip({path: zipPath}).then(function(zipFile) {
        return zipFile.read("hello", "utf8");
    }).then(function(contents) {
        assert.equal(contents, "Hello world\n");
    });
});

test('unzip can open Buffer', function() {
    var zipPath = path.join(__dirname, "test-data/hello.zip");
    return fs.readFile(zipPath)
        .then(function(buffer) {
            return unzip.openZip({buffer: buffer});
        })
        .then(function(zipFile) {
            return zipFile.read("hello", "utf8");
        })
        .then(function(contents) {
            assert.equal(contents, "Hello world\n");
        });
});
