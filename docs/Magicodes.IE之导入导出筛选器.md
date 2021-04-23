# Magicodes.IE之导入导出筛选器

![总体设计](https://docs.xin-lai.com/assets/Magicodes.IE.png)

[Magicodes.IE](https://github.com/dotnetcore/Magicodes.IE)是一个导入导出通用库，支持Dto导入导出以及动态导出，支持Excel、Word、Pdf、Csv和Html。在本篇教程，笔者将讲述如何使用Magicodes.IE的导入导出筛选器。在开始之前，我们需要先了解Magicodes.IE目前支持的筛选器：

| 接口                  | 说明                                                         |
| --------------------- | ------------------------------------------------------------ |
| IImportResultFilter   | 导入结果筛选器，可以修改导入结果包括验证错误信息（比如动态修改错误标注） |
| IImportHeaderFilter   | 导入列头筛选器，可以修改列名、值映射集合等等                 |
| IExporterHeaderFilter | 导出列头筛选器，可以修改列头、索引、值映射等等               |

## 导入结果筛选器（IImportResultFilter）的使用

导入结果筛选器可以修改导入结果包括验证错误信息（比如动态修改错误标注），非常适合对导入数据和错误验证内容进行二次动态加工，比如加入自定义校验逻辑、验证消息多语言翻译等等。接下来我们开始实战：

### 准备导入文件

如下图所示，我们准备了如下Excel导入文件：

![导入文件](https://docs.xin-lai.com/assets/image-20200921163108561.png)

[下载地址](https://github.com/dotnetcore/Magicodes.IE/blob/master/src/Magicodes.ExporterAndImporter.Tests/TestFiles/Errors/%E6%95%B0%E6%8D%AE%E9%94%99%E8%AF%AF.xlsx)

### 准备Dto

Excel准备好了，我们需要准备一个Dto：

```csharp
[ExcelImporter(ImportResultFilter = typeof(ImportResultFilterTest), IsLabelingError = true)]
public class ImportResultFilterDataDto1
{
    /// <summary>
    ///     产品名称
    /// </summary>
    [ImporterHeader(Name = "产品名称")]
    public string Name { get; set; }

    /// <summary>
    ///     产品代码
    ///     长度验证
    ///     重复验证
    /// </summary>
    [ImporterHeader(Name = "产品代码", Description = "最大长度为20", AutoTrim = false, IsAllowRepeat = false)]
    public string Code { get; set; }
}
```
如上述代码所示，我们创建了名为“ImportResultFilterDataDto1”的Dto，使用ExcelImporter特性中的ImportResultFilter属性指定了导入结果筛选器的类型。

### 创建类并实现接口IImportResultFilter

接下来我们就创建一个类并实现IImportResultFilter接口：

```csharp
public class ImportResultFilterTest : IImportResultFilter
{
    /// <summary>
    /// 本示例修改数据错误验证结果，可用于多语言等场景
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="importResult"></param>
    /// <returns></returns>
    public ImportResult<T> Filter<T>(ImportResult<T> importResult) where T : class, new()
    {
        var errorRows = new List<int>()
        {
            5,6
        };
        var items = importResult.RowErrors.Where(p => errorRows.Contains(p.RowIndex)).ToList();

        for (int i = 0; i < items.Count; i++)
        {
            for (int j = 0; j < items[i].FieldErrors.Keys.Count; j++)
            {
                var key = items[i].FieldErrors.Keys.ElementAt(j);
                var value = items[i].FieldErrors[key];
                items[i].FieldErrors[key] = value?.Replace("存在数据重复，请检查！所在行：", "Duplicate data exists, please check! Where:");
            }
        }
        return importResult;
    }
}
```
如上述代码所示，我们将重复错误的验证提示修改为了“Duplicate data exists, please check! Where”。接下来，我们需要编写导入代码：

### 编写导入代码

```csharp
    public async Task ImportResultFilter_Test()
    {
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Errors", "数据错误.xlsx");
        var labelingFilePath = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(ImportResultFilter_Test)}.xlsx");
        var result = await Importer.Import<ImportResultFilterDataDto1>(filePath, labelingFilePath);
    }
```
打开上述代码所示的标注文件路径，就可以看到验证提示被我们改成了英文：

![验证提示](https://docs.xin-lai.com/assets/image-20200918140908848.png)



## 导入列头筛选器（IImportHeaderFilter）的使用

导入列头筛选器可以修改列名、验证属性、值映射集合等等，非常适合动态修改列名、验证逻辑、值映射等等。和前面的一样，我们先得准备一个导入文件。

### 准备导入文件

![导入文件](https://docs.xin-lai.com/assets/image-20200921163656466.png)

[下载地址](https://github.com/dotnetcore/Magicodes.IE/blob/master/src/Magicodes.ExporterAndImporter.Tests/TestFiles/Import/%E5%AF%BC%E5%85%A5%E5%88%97%E5%A4%B4%E7%AD%9B%E9%80%89%E5%99%A8%E6%B5%8B%E8%AF%95.xlsx)

### 准备Dto

```csharp
/// <summary>
/// 导入学生数据Dto
/// IsLabelingError：是否标注数据错误
/// </summary>
[ExcelImporter(IsLabelingError = true, ImportHeaderFilter = typeof(ImportHeaderFilterTest))]
public class ImportHeaderFilterDataDto1
{
    /// <summary>
    ///     姓名
    /// </summary>
    [ImporterHeader(Name = "姓名", Author = "雪雁")]
    [Required(ErrorMessage = "学生姓名不能为空")]
    [MaxLength(50, ErrorMessage = "名称字数超出最大限制,请修改!")]
    public string Name { get; set; }

    /// <summary>
    ///     性别
    /// </summary>
    [ImporterHeader(Name = "性别")]
    [Required(ErrorMessage = "性别不能为空")]
    public Genders Gender { get; set; }

}
```
如上述代码所示，我们通过ImportHeaderFilter属性指定了列头筛选器类型。接下来，我们需要完成相关实现：

### 创建类并实现接口IImportHeaderFilter

```csharp
/// <summary>
/// 导入列头筛选器测试
/// 1）测试修改列头
/// 2）测试修改值映射
/// </summary>
public class ImportHeaderFilterTest : IImportHeaderFilter
{
    public List<ImporterHeaderInfo> Filter(List<ImporterHeaderInfo> importerHeaderInfos)
    {
        foreach (var item in importerHeaderInfos)
        {
            if (item.PropertyName == "Name")
            {
                item.Header.Name = "Student";
            }
            else if (item.PropertyName == "Gender")
            {
                item.MappingValues = new Dictionary<string, dynamic>()
                {
                    {"男",0 },
                    {"女",1 }
                };
            }
        }
        return importerHeaderInfos;
    }
}
```
通过上述代码，我们编写了一些测试：

1. 实现了IImportHeaderFilter
2. 将属性名称为“Name”的列的列头修改为“Student”
3. 将属性名称为“Gender”的列的列映射改为男女映射

接下来我们继续编写导入逻辑：

        public async Task ImportHeaderFilter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "导入列头筛选器测试.xlsx");
            var import = await Importer.Import<ImportHeaderFilterDataDto1>(filePath);
        }
如下图所示，我们成功的将Excel列名为“Student”的列导入到了Dto的Name属性，同时将男女转换为了枚举：

![验证](https://docs.xin-lai.com/assets/image-20200921164132097.png)

## 导出列头筛选器（IExporterHeaderFilter）的使用

导出列头筛选器可以修改列头、索引、值映射，非常适合动态修改导出逻辑，比如列头的中英转换，值映射动态逻辑等等。接下来我们一起来实战：

### 准备Dto并编写导出代码

```csharp
[Exporter(Name = "测试", TableStyle = "Light10", ExporterHeaderFilter = typeof(TestExporterHeaderFilter1))]
public class ExporterHeaderFilterTestData1
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
}



```
如上述Dto代码所示，我们通过导出特性Exporter的ExporterHeaderFilter属性指定了导出列头筛选器。

### 实现筛选器IExporterHeaderFilter

```csharp
public class TestExporterHeaderFilter1 : IExporterHeaderFilter
{
    /// <summary>
    /// 表头筛选器（修改名称）
    /// </summary>
    /// <param name="exporterHeaderInfo"></param>
    /// <returns></returns>
    public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
    {
        if (exporterHeaderInfo.DisplayName.Equals("名称"))
        {
            exporterHeaderInfo.DisplayName = "name";
        }
        return exporterHeaderInfo;
    }
}
```
如上述代码所示，我们实现了导出筛选器，并将显示名为“名称”的列修改为了“name”。

### 编写导出逻辑

```csharp
//导出
IExporter exporter = new ExcelExporter();
//使用GenFu生成测试数据
var data1 = GenFu.GenFu.ListOf<ExporterHeaderFilterTestData1>();
var result = await exporter.Export(filePath, data1);
```

使用上述代码导出后，我们来验证导出结果：

![导出结果](https://docs.xin-lai.com/assets/image-20200921165235468.png)

是不是So easy呢？当然我们还可以做一些其他的事情，比如修改忽略列：

```csharp
public class TestExporterHeaderFilter2 : IExporterHeaderFilter
{
    /// <summary>
    /// 表头筛选器（修改忽略列）
    /// </summary>
    /// <param name="exporterHeaderInfo"></param>
    /// <returns></returns>
    public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
    {
        if (exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore)
        {
            exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore = false;
        }
        return exporterHeaderInfo;
    }
}
```
## 如何使用容器注入筛选器

筛选器主要是为了满足大家能够在导入导出时支持动态处理，比如值映射等等。但是通过特性指定筛选器的话，那么如何支持依赖注入呢？不要慌，针对这个场景，我们也有考虑。

### 在ASP.NET Core的启动类（StartUp）注册容器

参考代码如下：

    public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
    {
    	AppDependencyResolver.Init(app.ApplicationServices);
        //添加注入关系
        services.AddSingleton<IImportResultFilter, ImportResultFilterTest>();
        services.AddSingleton<IImportHeaderFilter, ImportHeaderFilterTest>();
        services.AddSingleton<IExporterHeaderFilter, TestExporterHeaderFilter1>();	
    }
然后就尽情使用吧。值得注意的是：

1. **注入的筛选器类型的优先级高于特性指定的筛选器类型**，也就是当两者并存时，优先会使用注入的筛选器
2. 注入的筛选器是**全局的**，当注入多种类型的筛选器时，均会执行，接下来我们还会支持更多细节控制
3. 如果某个逻辑需要禁用所有筛选器，请参考下面部分
4. 此功能需要**2.4.0-beta2**或以上版本才支持

## 使用IsDisableAllFilter属性禁用所有的筛选器

如果某段导入导出需要禁用所有的筛选器，我们该如何处理？仅需将IsDisableAllFilter设置为true即可。导入导出特性均已支持。

