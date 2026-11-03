# excel-batch-core

纯 JavaScript 的 Excel 批量处理核心库:多表合并、按列拆分、xlsx/csv 格式互转。浏览器与 Node 通用,全程本地计算,数据不出设备。

在线工具版(开箱即用,表格不上传):**[excel.nightship.cc](https://excel.nightship.cc/?utm_source=github)**

## 功能

| 函数 | 作用 |
|---|---|
| `parseWorkbook(bytes)` | 从字节解析工作簿 → `{ sheetNames, sheets: {name: aoa} }` |
| `aoaToXlsxBytes(aoa)` | aoa → xlsx 字节 |
| `aoaToCsv(aoa)` | aoa → csv 文本(带 BOM,Excel 打开中文不乱码) |
| `csvToAoa(text)` | csv 文本 → aoa |
| `mergeSheets(workbooks, addSourceColumn)` | 多表纵向合并,可选插入"来源文件"列 |
| `splitByColumn(aoa, colIndex)` | 按列值分组拆分 → `{ 组名: aoa(含表头) }` |
| `multiSheetXlsxBytes(parts)` | 多个 aoa → 单个多 sheet xlsx 字节 |

细节行为:合并自动跳过全空行、短行补齐;拆分空值归入"(空白)"、组名截断 28 字符;sheet 名非法字符替换、超长截断、重名自动加后缀。

## 使用

依赖 [SheetJS](https://sheetjs.com/)(`npm install xlsx` 或浏览器引入 `xlsx.full.min.js`),以注入方式传入,不捆绑。

### Node

```js
const EXBCore = require("excel-batch-core"); // 自动 require("xlsx")
const fs = require("fs");

const wb = EXBCore.parseWorkbook(fs.readFileSync("a.xlsx"));
const merged = EXBCore.mergeSheets(
  Object.keys(wb.sheets).map(name => ({ name, firstAoa: wb.sheets[name] })),
  true, // 首列插入"来源文件"
);
fs.writeFileSync("merged.csv", Buffer.from(EXBCore.aoaToCsv(merged), "utf-8"));
```

### 浏览器

```html
<script src="xlsx.full.min.js"></script>
<script src="exb-core.js"></script>
<script>
  // window.EXBCore 可用
  const groups = EXBCore.splitByColumn(EXBCore.parseWorkbook(bytes).sheets["Sheet1"], 0);
</script>
```

## 测试

```bash
npm install xlsx
npm test
```

## 为什么本地处理

工资表、客户名单、订单数据这类敏感表格,上传到任何第三方服务器都有泄露与合规风险。这个库的全部计算发生在调用方进程内——这也是 [excel.nightship.cc](https://excel.nightship.cc/?utm_source=github) 的核心设计:纯前端,文件不出浏览器。

## License

[MIT](./LICENSE)
