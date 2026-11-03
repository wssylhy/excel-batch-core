/* excel-batch-core — 纯前端 Excel 批量合并/拆分/格式互转核心库
 * 浏览器与 Node 通用,本地处理,数据不出设备。
 * 在线工具版: https://excel.nightship.cc
 * 依赖 SheetJS (xlsx),以参数注入,不打包。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    var XLSX = null;
    try { XLSX = require("xlsx"); } catch (e) { /* 由调用方注入 */ }
    module.exports = factory(XLSX);
  } else {
    root.EXBCore = factory(root.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (XLSX) {
  "use strict";
  if (!XLSX) {
    throw new Error("excel-batch-core 需要 SheetJS: 浏览器请先引入 xlsx.full.min.js (window.XLSX),Node 请 npm install xlsx 或自行注入");
  }

  /* 从字节解析工作簿 → { sheetNames, sheets: {name: aoa} } */
  function parseWorkbook(bytes) {
    var wb = XLSX.read(bytes, { type: "array", cellDates: true, codepage: 936 });
    var out = { sheetNames: [], sheets: {} };
    wb.SheetNames.forEach(function (name) {
      out.sheetNames.push(name);
      out.sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: "" });
    });
    return out;
  }

  /* aoa → workbook bytes(xlsx) */
  function aoaToXlsxBytes(aoa) {
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    return XLSX.write(wb, { bookType: "xlsx", type: "array" }); // Uint8Array
  }

  /* aoa → csv 文本(带 BOM 头,Excel 直接打开中文不乱码) */
  function aoaToCsv(aoa) {
    return "\uFEFF" + XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(aoa));
  }

  /* csv 文本 → aoa */
  function csvToAoa(text) {
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    var wb = XLSX.read(text, { type: "string", raw: false });
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: "" });
  }

  /* 合并: workbooks = [{name, firstAoa}],addSourceColumn 为真时首列插入"来源文件" */
  function mergeSheets(workbooks, addSourceColumn) {
    if (!workbooks.length) return [];
    var header = workbooks[0].firstAoa.length ? workbooks[0].firstAoa[0].slice() : [];
    var rows = [];
    workbooks.forEach(function (wb) {
      var body = wb.firstAoa.slice(1);
      body.forEach(function (r) {
        if (String(r.join("")) === "") return; // 跳过全空行
        var row = r.slice();
        while (row.length < header.length) row.push("");
        if (addSourceColumn) row.unshift(wb.name);
        rows.push(row);
      });
    });
    var outHeader = addSourceColumn ? ["来源文件"].concat(header) : header;
    return [outHeader].concat(rows);
  }

  /* 按列拆分: aoa 按第 colIndex 列的值分组 → { 组名: aoa(含表头) };空值归入"(空白)" */
  function splitByColumn(aoa, colIndex) {
    var groups = {};
    var order = [];
    if (!aoa.length) return groups;
    var header = aoa[0];
    for (var i = 1; i < aoa.length; i++) {
      var row = aoa[i];
      if (String(row.join("")) === "") continue;
      var key = (row[colIndex] == null || String(row[colIndex]) === "") ? "(空白)" : String(row[colIndex]).slice(0, 28);
      if (!groups[key]) { groups[key] = [header]; order.push(key); }
      groups[key].push(row);
    }
    var stable = {};
    order.sort().forEach(function (k) { stable[k] = groups[k]; });
    return stable;
  }

  /* 多个 {sheetName, aoa} → 单个 xlsx 字节(多 sheet,sheet 名安全化去重) */
  function multiSheetXlsxBytes(parts) {
    var wb = XLSX.utils.book_new();
    var used = {};
    parts.forEach(function (p) {
      var name = String(p.sheetName || "Sheet").replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 31) || "Sheet";
      if (used[name]) { name = (name + "_" + used[name]).slice(0, 31); }
      used[name] = (used[name] || 0) + 1;
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(p.aoa), name);
    });
    return XLSX.write(wb, { bookType: "xlsx", type: "array" });
  }

  return {
    parseWorkbook: parseWorkbook,
    aoaToXlsxBytes: aoaToXlsxBytes,
    aoaToCsv: aoaToCsv,
    csvToAoa: csvToAoa,
    mergeSheets: mergeSheets,
    splitByColumn: splitByColumn,
    multiSheetXlsxBytes: multiSheetXlsxBytes,
  };
});
