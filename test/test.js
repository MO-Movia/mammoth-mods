import {describe, it} from "mocha";

export default function test(name, func) {
    if (typeof func === "object" && func !== null) {
        describe(name, function() {
            Object.keys(func).forEach(function(key) {
                test(key, func[key]);
            });
        });
    } else {
        it(name, func);
    }
}
