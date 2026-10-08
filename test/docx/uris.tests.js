import test from "../test.js";
import assert from "assert";
import * as zipfile from "../../lib/docx/uris.js";
test("uriToZipEntryName", {
    "when path does not have leading slash then path is resolved relative to base": function() {
        assert.equal(
            zipfile.uriToZipEntryName("one/two", "three/four"),
            "one/two/three/four"
        );
    },

    "when path has leading slash then base is ignored": function() {
        assert.equal(
            zipfile.uriToZipEntryName("one/two", "/three/four"),
            "three/four"
        );
    }
});
