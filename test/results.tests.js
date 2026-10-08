import test from "./test.js";
import assert from "assert";
import * as results from "../lib/results.js";
var Result = results.Result;

test("Result.combine removes any duplicate messages", function() {
    var first = new Result(null, [results.warning("Warning...")]);
    var second = new Result(null, [results.warning("Warning...")]);

    var combined = Result.combine([first, second]);

    assert.deepEqual(combined.messages, [results.warning("Warning...")]);
});
