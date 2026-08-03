"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

require(path.join(__dirname, "..", "xlsx-export.js"));

const bytes = globalThis.QCQuizExcel.buildWorkbook([
  {
    name: "測驗紀錄",
    headers: ["完成時間", "模式", "分數"],
    rows: [[new Date(2026, 7, 3, 20, 0, 0), "隨機測驗", 90]]
  },
  {
    name: "作答明細",
    headers: ["題號", "題目", "結果"],
    rows: [[1, "測試題目", "正確"]]
  }
]);

assert.equal(bytes[0], 0x50);
assert.equal(bytes[1], 0x4B);
const decoded = new TextDecoder().decode(bytes);
assert.match(decoded, /xl\/worksheets\/sheet1\.xml/);
assert.match(decoded, /測驗紀錄/);
assert.match(decoded, /作答明細/);

if (process.argv[2]) {
  fs.writeFileSync(process.argv[2], bytes);
}

console.log("xlsx export test OK");
