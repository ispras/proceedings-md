"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const child_process_1 = require("child_process");
function pandoc(src, args) {
    return new Promise((resolve, reject) => {
        let stdout = "";
        let stderr = "";
        let pandocProcess = (0, child_process_1.spawn)('pandoc', args);
        pandocProcess.stdin.end(src, 'utf-8');
        pandocProcess.stdout.on('data', (data) => {
            stdout += data;
        });
        pandocProcess.stderr.on('data', (data) => {
            stderr += data;
        });
        pandocProcess.on('exit', function (code) {
            if (stderr.length) {
                console.error("There was some pandoc warnings along the way:");
                console.error(stderr);
            }
            if (code == 0) {
                resolve(stdout);
            }
            else {
                reject(new Error("Pandoc returned non-zero exit code"));
            }
        });
    });
}
exports.default = pandoc;
//# sourceMappingURL=pandoc.js.map