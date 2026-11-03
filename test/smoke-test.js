/* 冒烟测试: node test/smoke-test.js (需先 npm install xlsx) */
"use strict";
const EXBCore = require("../exb-core.js");

function assert(cond, msg) {
  if (!cond) { console.error("FAIL: " + msg); process.exit(1); }
  console.log("ok - " + msg);
}

const header = ["姓名", "部门", "金额"];
const a = { name: "a.xlsx", firstAoa: [header, ["张三", "销售", 100], ["李四", "销售", 200]] };
const b = { name: "b.xlsx", firstAoa: [header, ["王五", "人事", 300]] };

// 合并 + 来源列
const merged = EXBCore.mergeSheets([a, b], true);
assert(merged.length === 4, "合并行数 = 表头 + 3 数据行");
assert(merged[1][0] === "a.xlsx" && merged[3][0] === "b.xlsx", "来源列标注正确");
assert(merged[0][0] === "来源文件", "表头首列为来源文件");

// 按列拆分
const noSrc = EXBCore.mergeSheets([a, b], false);
const groups = EXBCore.splitByColumn(noSrc, 1);
assert(Object.keys(groups).sort().join() === "人事,销售", "拆分出 2 个组");
assert(groups["销售"].length === 3, "销售组 = 表头 + 2 行");

// xlsx 字节往返
const bytes = EXBCore.aoaToXlsxBytes(noSrc);
const back = EXBCore.parseWorkbook(bytes);
assert(back.sheets["Sheet1"].length === 4, "xlsx 往返行数一致");

// csv 带 BOM 且可逆
const csv = EXBCore.aoaToCsv(noSrc);
assert(csv.charCodeAt(0) === 0xfeff, "csv 带 BOM 头");
const backAoa = EXBCore.csvToAoa(csv);
assert(backAoa.length === 4 && backAoa[1][0] === "张三", "csv 往返一致");

// 多 sheet
const multi = EXBCore.multiSheetXlsxBytes([
  { sheetName: "销售", aoa: a.firstAoa },
  { sheetName: "人事", aoa: b.firstAoa },
  { sheetName: "人事", aoa: b.firstAoa }, // 重名 → 人事_1
]);
const back2 = EXBCore.parseWorkbook(multi);
assert(back2.sheetNames.join() === "销售,人事,人事_1", "多 sheet 名去重正确");

console.log("\n全部通过 ✓");
