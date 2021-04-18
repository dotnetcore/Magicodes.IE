# Magicodes.IE之花式导出

![总体设计](https://docs.xin-lai.com/assets/Magicodes.IE.png)

[Magicodes.IE](https://github.com/dotnetcore/Magicodes.IE)是一个导入导出通用库，支持Dto导入导出以及动态导出，支持Excel、Word、Pdf、Csv和Html。在本篇教程，笔者将讲述如何使用Magicodes.IE进行花式导出。

在本篇教程，笔者主要讲述如何使用IE进行花式导出并满足客户爸爸的需求。

## 同一个数据源拆分Sheet导出

通常情况下，客户爸爸的需求是比较正常的，比如在数据量大时，希望将数据进行拆分导出。

这时候我们就需要使用IE按部就班开发了，先创建Dto：

```csharp
[ExcelExporter(Name = "测试2", TableStyle = "None", AutoFitAllColumn = true, MaxRowNumberOnASheet = 100)]
public class ExportTestDataWithSplitSheet
{
    [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
    public string Text { get; set; }

    [ExporterHeader(DisplayName = "普通文本")] public string Text2 { get; set; }

    [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
    public string Text3 { get; set; }

    [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
    public decimal Number { get; set; }

    [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
    public string Name { get; set; }

    /// <summary>
    /// 时间测试
    /// </summary>
    [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
    public DateTime Time1 { get; set; }

    /// <summary>
    /// 时间测试
    /// </summary>
    [ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
    public DateTime? Time2 { get; set; }

    public DateTime Time3 { get; set; }

    public DateTime Time4 { get; set; }
}
```
如上述Dto定义所示，我们通过**MaxRowNumberOnASheet**属性指定了**每个Sheet最大的行数**，接下来仅需使用普通导出即可自动拆分Sheet导出：

```csharp
        var result = await exporter.Export(filePath,
            GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(300));
```
![数据拆分Sheet导出](https://docs.xin-lai.com/assets/image-20200927105735213.png)

是不是非常简单？作为一个正直和诚实的人，这时候我们可以评估为2天的工作量。

## 多个数据源多Sheet导出

过了一段时间，客户爸爸厌倦了各种表格，他有一个残暴的想法——乙方渣渣，能不能把这个表格做成一个表格导出！为了不被甲方爸爸按在地上摩擦，我们先跪下来。在各种讨价还价之后，我们Get到了5天的工作量。

对于导出多个数据，IE也做了充分的考虑：

**Dto1:**

```csharp
[ExcelExporter(Name = "测试", TableStyle = "Light10", AutoFitAllColumn = true, AutoFitMaxRows = 5000)]
public class ExportTestDataWithAttrs
{
    [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
    public string Text { get; set; }
    [ExporterHeader(DisplayName = "普通文本")] public string Text2 { get; set; }
    [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
    public string Text3 { get; set; }
    [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
    public decimal Number { get; set; }
    [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
    public string Name { get; set; }

    /// <summary>
    /// 时间测试
    /// </summary>
    [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
    public DateTime Time1 { get; set; }

    /// <summary>
    /// 时间测试
    /// </summary>
    [ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
    public DateTime? Time2 { get; set; }

    [ExporterHeader(Width = 100)]
    public DateTime Time3 { get; set; }

    public DateTime Time4 { get; set; }

    /// <summary>
    /// 长数值测试
    /// </summary>
    [ExporterHeader(DisplayName = "长数值", Format = "#,##0")]
    public long LongNo { get; set; }
}
```
**Dto2:**

```csharp
[ExcelExporter(Name = "测试2", TableStyle = "None", AutoFitAllColumn = true, MaxRowNumberOnASheet = 100)]
public class ExportTestDataWithSplitSheet
{
    [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
    public string Text { get; set; }

    [ExporterHeader(DisplayName = "普通文本")] public string Text2 { get; set; }

    [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
    public string Text3 { get; set; }

    [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
    public decimal Number { get; set; }

    [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
    public string Name { get; set; }

    /// <summary>
    /// 时间测试
    /// </summary>
    [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
    public DateTime Time1 { get; set; }

    /// <summary>
    /// 时间测试
    /// </summary>
    [ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
    public DateTime? Time2 { get; set; }

    public DateTime Time3 { get; set; }

    public DateTime Time4 { get; set; }
}
```
以上代码定义了2个Dto，大家可以根据实际情况准备更多。接下来我们利用开篇所说的API来进行花式导出：

```csharp
        var list1 = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();
        var list2 = GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(30);
        var result = await exporter
            .Append(list1)
            .SeparateByColumn().Append(list2)
            .SeparateByColumn().Append(list2)
            .ExportAppendData(filePath);
```
如上述代码所示，我们通过**Append**添加了三个数据源，通过两次**SeparateByColumn**进行了拆分，最后通过**ExportAppendData**来导出：

![通过列拆分导出](https://docs.xin-lai.com/assets/image-20200927142653458.png)

导出结果如图所示。值得注意的是，两个Dto使用了不同的主题，因此多个导出也保持了不同的导出风格，是不是很赞呢？客户爸爸也觉得很赞，但是他还是觉得应该按Sheet拆分会比较直观。于是你被乙方领导按在地上摩擦了一会，然后给了你两天的时间。

## 多个数据源按Sheet拆分导出

因为摩擦产生了静电，所以你很快想到了IE肯定有相关的实现：

```csharp
var result = exporter
    .Append(list1, "sheet1")
    .SeparateBySheet()
    .Append(list2)
    .ExportAppendData(filePath);
```

如上述代码所示，我们将分割函数改为了SeparateBySheet，结果如下图所示：

![多个数据源按Sheet拆分导出](https://docs.xin-lai.com/assets/image-20200927140130735.png)

不过值得注意的是，Append函数支持传递Sheet名称来覆盖默认的Sheet命名，以便大家可以通过这些API动态灵活的导出。

## 多个数据源按行拆分导出

客户爸爸收到了你的更改，很是开心，决定给你一个奖赏——这不是我要的，我要分行导出。在被摩擦的几十年生涯中，你深刻的知道怼怒的结果无法是被一次一次的摩擦。

不过这次你心里有数，默默的报了7天的工作量，使用IE秒改，然后花了7天的时间来演戏：

```csharp
var result = await exporter
    .Append(list1)
    .SeparateByRow()
    .Append(list2)
    .ExportAppendData(filePath);
```

如上述代码所示，在导出领域，IE不是万能的，但是没有IE是万万不能的。通过修改**SeparateByRow**，我们就毫秒级完成了客户的需求：

![多个数据源按行拆分导出](https://docs.xin-lai.com/assets/image-20200927140422201.png)

7天后，客户拿到报表，欣喜之余习惯性的又想摩擦，哦，指出了一个问题：数据量太大，我希望表头时时刻刻的展现在我眼前！然后你装作苦逼的再报了7天的工作量，再次祭出IE秒改：

```csharp
var result = await exporter
    .Append(list1)
    .SeparateByRow()
    .AppendHeaders()
    .Append(list2)
    .ExportAppendData(filePath);
```

如上述代码所示，我们通过AppendHeaders完成了追加表头的需求，从此走上了人生巅峰：

![追加表头](https://docs.xin-lai.com/assets/image-20200927140527365.png)

## 最后

通过本篇教程，我想大家明白了一个道理：人生如戏，全靠演技。当你有IE作为后盾时，在甲方爸爸面前，你就可以尽情的跪拜了！

不过我们还是来做一个总结，在本教程中，只要你掌握了以下API，你就可以赢取白富美，走上人生巅峰了：

| API              | 说明                          |
| ---------------- | ----------------------------- |
| Append           | 追加数据源，支持传递Sheet名称 |
| AppendHeaders    | 追加表头                      |
| SeparateByColumn | 通过追加Column分割导出        |
| SeparateBySheet  | 通过Sheet分割导出             |
| SeparateByRow    | 通过追加行来分割导出          |
| ExportAppendData | 导出追加数据                  |

