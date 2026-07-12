# 迁移指南:Magicodes.IE.Excel → Magicodes.IE.IO

本文档帮助把现有 `Magicodes.IE.Excel`（基于 EPPlus）迁移到新的 `Magicodes.IE.IO`。两个包可以共存，逐步迁移。

---

## 为什么迁移

| 维度 | 老的 `Magicodes.IE.Excel` | 新的 `Magicodes.IE.IO` |
|---|---|---|
| EPPlus 依赖 | vendored(嵌 100+MB) | 运行时无；`netstandard2.0` 带兼容性依赖 |
| 性能 | 中(EPPlus DOM 模型) | 更轻,但要按场景看基准 |
| AOT / Trim | 不友好(EPPlus 反射) | 普通 .NET 应用无需额外配置；NativeAOT/Trim 场景为 DTO 添加 `[XlsxExportable]` 即启用已验证的生成读写路径（包含自定义 `CellConverter<T>`） |
| 多 sheet | 复杂 API | `WriteWorkbookToBytes(...)` 一行 |
| 异步流 | 无 | `IAsyncEnumerable<T>` |
| 模板导出 | 有(复杂) | `Xlsx.ExportByTemplateAsync` |
| Reader | 有 | 有(`Xlsx.Read<T>` / `Xlsx.ReadAsync<T>`) |
| 协议 | MIT | MIT |

---

## 安装

```bash
# 保留老的 Magicodes.IE.Excel 同时引入新的
dotnet add package Magicodes.IE.IO
```

目标框架:`netstandard2.0` / `net6.0` / `net8.0` / `net10.0`。

---

## API 映射(导出)

### 1. 简单导出(零配置)

老写法:

```csharp
public class OrderDto
{
    [Exporter(Name = "订单号")]
    public string OrderNo { get; set; }
    public decimal Amount { get; set; }
}

await new ExcelExporter().Export("/path/to/file.xlsx", data);
```

新写法:

```csharp
public class OrderDto
{
    [ExporterHeader(Name = "订单号")]
    public string OrderNo { get; set; }
    public decimal Amount { get; set; }
}

Xlsx.Write("/path/to/file.xlsx", data);
```

要点:

- `ExcelExporter` → `Xlsx` 静态入口
- `Export(string path, data)` → `Xlsx.Write(path, data)`
- `ExporterAttribute` → `ExporterHeaderAttribute`(老 attribute 在 IE.Core 里仍可用,但推荐迁移)

### 2. 字节数组导出

老:

```csharp
var bytes = await new ExcelExporter().ExportAsByteArray(data);
```

新:

```csharp
var bytes = Xlsx.ToBytes(data);
```

### 3. 列映射(老 `ExportDto` / 新 `ExportProfile<T>`)

老:

```csharp
[ExcelExporter(Name = "订单导出", HeaderFontColor = "FF0000")]
public class OrderExportDto : ExportDto<Order> { }
```

新:

```csharp
var profile = new ExportProfile<Order>()
    .Sheet("订单表")
    .Column(x => x.OrderNo, c => c.WithName("订单号"));

var bytes = Xlsx.ToBytes(orders, profile);
```

要点:

- 老的 `ExportDto` 抽象 + 配置类 → 新的 `ExportProfile<T>` fluent builder
- 列属性改用 `.Column(x => x.Foo, c => c.WithName(...).WithFormat(...))`
- 老 `[ExporterHeader]` 在新包里继续兼容,推荐和 `[Display]` 语义对齐

### 4. 模板导出(老 `ExportByTemplate`)

老:

```csharp
await new ExcelExporter().ExportByTemplate("/template.xlsx", "/output.xlsx", data);
```

新:

```csharp
await Xlsx.ExportByTemplateAsync("/template.xlsx", "/output.xlsx", data);
```

模板语法兼容 `{{PropertyName}}` / `{{#List}}...{{/List}}`。

### 5. 读取(老 `Import` / 新 `Xlsx.Read`)

老:

```csharp
var orders = await new ExcelImporter().Import<OrderDto>(stream);
```

新:

```csharp
var orders = new List<OrderDto>();
await foreach (var order in Xlsx.ReadAsync<OrderDto>(stream))
    orders.Add(order);

// 或同步:
var syncOrders = Xlsx.Read<OrderDto>(stream).ToList();
```

要点:

- `Import<T>` → `Xlsx.Read<T>` / `Xlsx.ReadAsync<T>`
- `[ImportHeader(Name = "...")]` → `[ImporterHeader(Name = "...")]`
- 老 `ImportDtoBase` 配置类对应新 `XlsxReadOptions<T>`

### 6. 不再推荐库内 DI 抽象

老代码如果依赖库内的导入/导出抽象,建议直接改成调用 `Xlsx` 静态入口,然后在应用层自己包一层业务服务。

```csharp
public sealed class OrderReportWriter
{
    public void Write(string path, IEnumerable<OrderDto> data)
        => Xlsx.Write(path, data);
}
```

---

## 行为差异

1. **attribute 改名**:`[Exporter(Name = ...)]` → `[ExporterHeader(Name = ...)]`;`[ImportHeader(Name = ...)]` → `[ImporterHeader(Name = ...)]`
2. **特性优先级**:fluent `cfg.WithName() > [ExporterHeader(Name=)] > [Display(Name=)] > [Description] > 属性名`
3. **多 sheet**:`WriteWorkbookToBytes(sheet1, sheet2, ...)` 接受 `Sheet<T>` 对象
4. **图片导出**:`XlsxWriter.AddImage(byte[], ext, fromCell, toCell)`
5. **reader 类型**:`Xlsx.Read<T>` 约束 `T : new()`
6. **没有 `[ExportDto]` 抽象**:`ExportProfile<T>` fluent 替代;`ExportDtoBase` 改为组合 `ExportProfile<T>`
7. **没有 `Magicodes.ExporterAndImporter.AspNetCore` 包**:`Xlsx` 静态入口即推荐入口

---

## 不支持的功能

- **图表 / 宏**:xlsx 标准允许但未实现;若需要,继续用 EPPlus
- **PDF / Word / HTML / CSV / JSON 输出**:`Magicodes.IE.IO` 当前只做 `.xlsx`
- **Conditional Formatting 变体**:只支持基础规则,color scale / data bar / icon set 暂未实现
- **Pivot Table / 自定义 XML part**:未实现
- **大文件**:`Magicodes.IE.IO` 支持流式 `IAsyncEnumerable`，推荐用 `Xlsx.WriteAsync(stream, ...)`；但当前 ZIP writer 不支持 ZIP64，仍受 ZIP32 和单 worksheet 行列上限约束
- **Reader 范围**:`Xlsx.Read<T>` 读 cell 值 + SST;不读公式结果 / 条件格式 / 批注 / 命名范围

---

## 共存策略

两个包可同时引用,新代码用 `Magicodes.IE.IO`,老代码继续用 `Magicodes.IE.Excel`。逐步:

1. 新功能用新包
2. 读老表用 `Xlsx.Read`
3. 写老表用新 `Xlsx` 包
4. 测试通过后,移除 `Magicodes.IE.Excel` 引用
---

## 性能参考

最新 benchmark 的完整数字请直接看同仓库里的 `BenchmarkDotNet.Artifacts/results/` 报告。
