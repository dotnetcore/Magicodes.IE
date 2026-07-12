# Magicodes.IE.IO

Excel I/O 库，热点写入路径以低分配为目标。本文涵盖公共 API 速览、能力边界与选型建议。

覆盖常用 Excel 能力（数据验证 / 公式 / 行高 / SST / 表格 / 保护 / 打印 / 大纲 / 批注 / 命名范围 / 条件格式）。

---

## 公开 API 一览

高层入口都在 `Xlsx` 静态类中：

| 分组 | API | 说明 |
|---|---|---|
| 写 | `Write(path, data, cfg?, opt?)` | 写到文件 |
| 写 | `Write(Stream, data, cfg?, opt?)` | 写到流 |
| 写 | `Write(IBufferWriter<byte>, data, cfg?, opt?)` | 主低分配路径 |
| 写 | `WriteAsync(Stream, IAsyncEnumerable<T>, cfg?, opt?, ct)` | 异步枚举边查边写 |
| 写 | `WriteAsync(Stream, IEnumerable<T>, cfg?, opt?, ct)` | 已物化集合的异步流写入 |
| 写 | `ToBytes(data, cfg?, opt?)` | 便利层，返回 `byte[]` |
| 读 | `Read<T>(Stream, profile?, onError?)` | 同步读取 |
| 读 | `ReadAsync<T>(Stream, profile?, onError?, ct)` | 异步枚举读取 |
| 多 sheet | `WriteWorkbook(Stream, params Sheet[])` | 多 sheet 写流 |
| 多 sheet | `WriteWorkbook(IBufferWriter<byte>, params Sheet[])` | 多 sheet 低分配 |
| 多 sheet | `WriteWorkbookToBytes(params Sheet[])` | 多 sheet 便利层 |
| 模板 | `ExportByTemplateAsync(path, path, data)` | 模板导出(文件) |
| 模板 | `ExportByTemplateAsync(Stream, Stream, data)` | 模板导出(流) |

配置类型：`ExportProfile<T>`（列名/格式/忽略/sheet 名/fluent DSL）、`XlsxWriteOptions`（压缩/严格引用）、`XlsxReadOptions<T>`（列映射/自定义转换器）、`XlsxReadErrorInfo`（读取错误信息）和 `Sheet<T>` / `Sheet`（多 sheet 元素）。

选型建议：

- 写入文件、响应流或大数据时，优先使用 `Write(Stream, ...)` 或 `Write(IBufferWriter<byte>, ...)`。
- 数据源支持 `IAsyncEnumerable<T>`，或需要异步等待流 I/O 时，使用 `WriteAsync(...)`。
- 只需要一次性取得结果时，使用 `ToBytes(...)`；该便利层必然物化 `byte[]`。
- `Read(...)` / `ReadAsync(...)` 读取 xlsx；异步版本的收益在流 I/O 等待，单元格解析与对象映射仍是同步 CPU 工作。

写入 API 不拥有调用方提供的输出流，不会负责关闭它。直接使用底层 `XlsxWriter` 时，必须先添加至少一个 worksheet，再调用 `Complete()`。

---

## 安装

```bash
dotnet add package Magicodes.IE.IO
```

目标框架：`netstandard2.0` / `net6.0` / `net8.0` / `net10.0`。
`net6.0` 及以上目标仅依赖 BCL；`netstandard2.0` 目标由 NuGet 自动带入兼容性依赖。

---

## 快速上手

### 零配置

```csharp
Xlsx.Write("/tmp/orders.xlsx", orders);
```

表头 = 属性名，列序 = 声明序，自动 inline string / number / datetime / bool / enum / struct / record。

### fluent profile

```csharp
var bytes = Xlsx.ToBytes(orders, p => p
    .Sheet("订单表")
    .Column(x => x.OrderNo, c => c.WithName("订单号").WithWidth(30))
    .Column(x => x.Amount,  c => c.WithFormat("0.00"))
    .Ignore(x => x.CreatedAt)
    .WithFreezeHeader(true));
```

### 多 sheet

```csharp
var bytes = Xlsx.WriteWorkbookToBytes(
    new Sheet<Order>("Orders", orders),
    new Sheet<Item>("Items",  items));
```

不想要 `byte[]`：

```csharp
using var fs = File.Create("/tmp/report.xlsx");
Xlsx.WriteWorkbook(fs,
    new Sheet<Order>("Orders", orders),
    new Sheet<Item>("Items", items));
```

### 读取 xlsx

```csharp
// 同步
var rows = Xlsx.Read<Order>(stream).ToList();

// 异步
await foreach (var o in Xlsx.ReadAsync<Order>(stream)) { ... }
```

### 异步流写入（IAsyncEnumerable）

```csharp
async IAsyncEnumerable<Order> YieldOrders()
{
    await foreach (var o in dbContext.Orders.AsAsyncEnumerable())
        yield return o;
}

await Xlsx.WriteAsync(stream, YieldOrders());   // 边查边写，输出 I/O 异步等待
```

### 模板导出

基于一个 `.xlsx` 模板，把单元格里 `{{属性名}}` 占位符替换为数据值；`{{#集合}}…{{/集合}}` 列表块按集合逐行展开。模板原有的样式、合并、图片、公式都保留。

#### 基本用法

```csharp
// 文件 → 文件
await Xlsx.ExportByTemplateAsync("template.xlsx", "output.xlsx", data);

// 流 → 流（非可寻流自动 buffer 到内存，无需 seekable）
await Xlsx.ExportByTemplateAsync(templateStream, outputStream, data);
```

#### 单值占位符 `{{属性名}}`

模板单元格里写 `{{属性名}}`，导出时按属性名（大小写不敏感）反射取值替换。一个单元格内可放多个占位符加任意字面文本：

```csharp
public class TemplateHolder
{
    public string CustomerName { get; set; } = "";
    public string OrderNo { get; set; } = "";
    public List<TemplateCell> Items { get; set; } = new();
}

// 模板 A1 单元格：客户:{{CustomerName}}    A2 单元格：订单号:{{OrderNo}}
await Xlsx.ExportByTemplateAsync("tpl.xlsx", "out.xlsx",
    new TemplateHolder { CustomerName = "张三", OrderNo = "SO-1" });
// 结果：A1 = 客户:张三   A2 = 订单号:SO-1
```

行为细则：
- 值为 `null` → 替换为空串。
- 属性找不到 → 占位符**原样保留**（不报错，便于发现拼写错误）。
- 只支持 T 的**顶层属性**，不支持 `{{A.B}}` 嵌套路径。
- 值经 `Convert.ToString(value, InvariantCulture)` 转文本，并按上下文自动 XML 转义：在文本节点里转 `<`/`&` 等，在属性值里转引号。如 `CustomerName="A&B<C>"` → `A&amp;B&lt;C&gt;`。
- ⚠️ 日期/数字被替换为**不变区域性字符串**（如 `2026/07/01 00:00:00`、`12.5`），不会套用模板单元格的数字格式。如需保留日期/数字格式，自行先格式化为字符串再绑定。

#### 列表块 `{{#集合}}…{{/集合}}`

`{{#Items}}` 与 `{{/Items}}` 之间的内容作为"行模板"按 `Items` 集合逐项展开，块内 `{{字段}}` 取自当前项。**两个标记要放在 `<row>` 元素之间**（不在单元格里），中间是完整的一行或多行作为重复模板：

```csharp
public class TemplateCell
{
    public string Name { get; set; } = "";
    public int Qty { get; set; }
    public decimal Price { get; set; }
}
```

模板 sheet XML（节选）：

```xml
<row r="1"><c r="A1" t="inlineStr"><is><t>客户:{{CustomerName}}</t></is></c></row>
{{#Items}}<row r="1"><c r="A1" t="inlineStr"><is><t>{{Name}} x{{Qty}} ${{Price}}</t></is></c></row>{{/Items}}
```

`Items` 有 2 项（`{Name="A",Qty=2,Price=5}`、`{Name="B",Qty=3,Price=7}`）时，导出后行号自动平移：

```xml
<row r="1">…客户:李四…</row>
<row r="2"><c r="A2"…>A x2 $5</c></row>
<row r="3"><c r="A3"…>B x3 $7</c></row>
```

- `Items` 必须是 T 的 `IEnumerable` 属性。集合为空 → 块整段移除、块后行上移回收模板行位。
- 块内/块后的 `<row r="N">`、`ref=`/`location=` 属性、以及公式 `<f>` 里的 A1 引用都同步平移；公式**引号内的字符串原样不动**。
- **不支持嵌套**列表块（`{{#A}}…{{#B}}…{{/B}}…{{/A}}` 会被非贪婪正则错配）。

#### 工作表名 `{{!Sheet:Name=名称}}`

在 `workbook.xml` 的 sheet name 处写 `{{!Sheet:Name=销售明细}}`，导出时替换为"销售明细"（**静态**替换为占位符里写定的名称，不绑定数据）。

#### 处理范围与资源

- 只替换 `xl/worksheets/*`、`xl/sharedStrings.xml`、`xl/workbook.xml` 三类 part 里的占位符；样式、图片、图表等 part 不处理。
- 输入流不可寻时自动复制到内存再处理；输出流由调用方拥有，库不主动 Dispose（流重载）。

### 写入选项

```csharp
var bytes = Xlsx.ToBytes(orders, p => p.WithFreezeHeader(), new XlsxWriteOptions
{
    Compression = CompressionLevel.NoCompression,
    StrictCellReferences = false,
});
```

---

## 属性与配置

优先级：**fluent `cfg.WithName()` > `[ExporterHeader(Name=)]` > `[Display(Name=)]` > `[Description(...)]` > 属性名**

```csharp
public class Order
{
    [ExporterHeader(Name = "订单号", Width = 30)]
    public string OrderNo { get; set; }

    [DisplayFormat(DataFormatString = "0.00")]
    public decimal Amount { get; set; }

    [ExporterHeader(IsIgnore = true)]
    public DateTime CreatedAt { get; set; }
}

var bytes = Xlsx.ToBytes(orders);   // 表头 = 订单号 / Amount;CreatedAt 自动忽略
```

---

## 高级功能（默认隐藏）

下面是高级功能，不是普通导出的首选入口。

### 数据验证(Data Validation)

```csharp
using var ms = new MemoryStream();
using var writer = new XlsxWriter(ms, "订单表");
writer.AddDataValidation(new DataValidation("C2:C1000", DataValidationType.List, "\"已下单,已发货,已完成\""));
```

### 公式（列级 + 行级占位）

```csharp
var bytes = Xlsx.ToBytes(items, p => p
    .Column(x => x.Qty, c => c.WithName("数量"))
    .Column(x => x.Price, c => c.WithName("单价"))
    .Column(x => x.Total, c => c.WithName("合计").WithFormula("A{row}*B{row}")));
```

`{row}` 占位当前 1-based sheet 行号（自动展开为 `A2*B2` / `A3*B3` / ...）。

### Shared Strings Table（大文件去重）

```csharp
var bytes = Xlsx.ToBytes(orders, p => p.WithAutoSst(true));
```

`WithAutoSst(true)` 走启发式：同步和异步路径都会预扫最多前 64 行，字符串去重比例低于
70% 时切到 SST。`RowFilter` 会先执行，null 数据项不会参与探测；行数少于 16 行或使用
`NoCompression` 时保持 inline string。

### 表格 / 保护 / 打印 / 命名范围 / 批注 / 条件格式 / 大纲

`XlsxWriter.AddTable(...)` 添加 Excel 表格(写入 `xl/tables`)；表格样式用 `TableDefinition.WithTableStyle(TableStyle.Xxx)`(内置样式枚举,IntelliSense 可提示全部 60 个内置名)指定,或用 `WithTableStyle("自定义样式名")` 兜底。其余分别走 `SetSheetProtection(...)` / `SetPageSetup(...)` / `AddNamedRange(...)` / `AddComment(...)` / `AddConditionalFormatting(...)` / `SetOutline(...)`。详细示例见 `tests/Magicodes.IE.IO.Tests/`。

---

## 压缩档位

```csharp
// 默认 Fastest
Xlsx.ToBytes(data);

// NoCompression:大表上更快，但产物明显变大
Xlsx.ToBytes(data, options: new XlsxWriteOptions { Compression = CompressionLevel.NoCompression });
```

默认 `Fastest` 已覆盖绝大多数场景。若导出是高频热点、CPU 比带宽更紧张，用 `NoCompression` 换更快的写入（代价是文件变大）；若文件要经网络传输、体积优先，再考虑 `Optimal`。

---

## 性能

热点写入路径以低分配为目标：`Stream` / `IBufferWriter<byte>` 是主低分配路径，`ToBytes(...)` 便利层因需物化 `byte[]` 分配更高。具体基准数字见仓库内 BenchmarkDotNet 产物。严格来说，当前目标是“低分配主路径 + 便利层可用”，不是所有场景绝对零分配。

---

## 与老 Magicodes.IE.Excel 的区别

| | 老的 `Magicodes.IE.Excel` | 新的 `Magicodes.IE.IO` |
|---|---|---|
| 第三方依赖 | 有 | 运行时无；`netstandard2.0` 带兼容性依赖 |
| 性能 | 中（DOM 模型） | **流式低分配** |
| reader | 有 | `Read<T>` / `ReadAsync<T>` |
| 多 sheet | 复杂 | `WriteWorkbookToBytes(...)` 一行 |
| 异步流 | 无 | `IAsyncEnumerable` |
| AOT / Trim | 不友好 | 普通 .NET 应用无需额外配置；NativeAOT/Trim 场景为 DTO 添加 `[XlsxExportable]` 即启用已验证的生成读写路径（包含自定义 `CellConverter<T>`） |
| API 风格 | 老抽象层 / 多接口入口 | `Xlsx.Write` / `Xlsx.Read` 单一静态入口 |

**迁移**：把旧的导出/导入抽象入口替换为 `Xlsx.Write(...)` / `Xlsx.ToBytes(...)` / `Xlsx.Read(...)`，`ExporterHeaderAttribute` 仍然兼容，`ExportDtoAttribute` 不再需要（直接使用 `ExportProfile<T>`）。

---

## AOT 与裁剪

`[XlsxExportable]` 标记的 DTO 会由 Source Generator 生成属性 getter、cell writer、列元数据和 cell setter，读写路径都不需要访问 DTO 的反射元数据。未标记的类型在动态代码不可用时会回退到反射。

普通 .NET 应用可以直接使用 `Xlsx`，无需为 DTO 添加标记。需要 NativeAOT/Trim 时，推荐为 DTO 添加 `[XlsxExportable]`，启用与 `System.Text.Json` 配合 `JsonSerializerContext` 类似的生成元数据路径；该路径包含自定义 `CellConverter<T>`，且 `Read` 通过基类虚分发调用，不依赖运行时反射。

---

## 已知限制

- **API 形态**：`ToBytes(...)` 最终仍会物化 `byte[]`；真大数据请优先用 `Write(streamOrBuffer, ...)`
- **Reader**：`Xlsx.Read<T>` 不读公式结果（只读缓存值）；不读条件格式 / 批注 / 表格
- **ZIP**：当前 writer 不支持 ZIP64，超过 ZIP32 限制的工作簿需要使用其他工具链

---

## 协议

MIT