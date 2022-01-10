# Excel模板导出之动态导出

## 说明

目前Magicodes.IE已支持Excel模板导出时使用`JObject`、`Dictionary`和`ExpandoObject`来进行动态导出，具体使用请看本篇教程。

Tips:

- [ExpandoObject 类：表示可在运行时动态添加和删除其成员的对象](https://docs.microsoft.com/zh-cn/dotnet/api/system.dynamic.expandoobject?view=net-5.0&WT.mc_id=DT-MVP-5004079)

本功能的想法、部分实现初步源于[arik](https://gitee.com/arik)的贡献，这里再次感谢[arik](https://gitee.com/arik)！

在开始本篇教程之前，我们重温一下模板导出的语法：

```
  {{Company}}  //单元格渲染
  {{Table>>BookInfos|RowNo}} //表格渲染开始语法
  {{Remark|>>Table}}//表格渲染结束语法
  {{Image::ImageUrl?Width=50&Height=120&Alt=404}} //图片渲染
  {{Image::ImageUrl?w=50&h=120&Alt=404}} //图片渲染
  {{Image::ImageUrl?Alt=404}} //图片渲染
  {{Formula::AVERAGE?params=G4:G6}}  //公式渲染
  {{Formula::SUM?params=G4:G6&G4}}   //公式渲染
```

如果您对Magicodes.IE的模板导出不太了解，请阅读以下教程：

[Excel模板导出之导出教材订购表](9.Excel模板导出之导出教材订购表.md)

接下来，我们开始本篇教程：

## 1.安装包Magicodes.IE.Excel

```powershell
Install-Package Magicodes.IE.Excel
```

## 2.准备Excel模板文件

参考如图：

![模板文件](../res/image-20210308175620226.png)

该文件可以在测试工程中找到，文件名为【DynamicExportTpl.xlsx】。

## 3.使用JObject完成动态导出

代码比较简单，如下所示：

```csharp
    string json = @"{
      'Company': '雪雁',
      'Address': '湖南长沙',
      'Contact': '雪雁',
      'Tel': '136xxx',
      'BookInfos': [
        {'No':'a1','RowNo':1,'Name':'Docker+Kubernetes应用开发与快速上云','EditorInChief':'李文强','PublishingHouse':'机械工业出版社','Price':65,'PurchaseQuantity':10000,'Cover':'https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965','Remark':'备注'},
        {'No':'a1','RowNo':1,'Name':'Docker+Kubernetes应用开发与快速上云','EditorInChief':'李文强','PublishingHouse':'机械工业出版社','Price':65,'PurchaseQuantity':10000,'Cover':'https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965','Remark':'备注'}
      ]
    }";
    var jobj = JObject.Parse(json);
    //模板路径
    var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
        "DynamicExportTpl.xlsx");
    //创建Excel导出对象
    IExportFileByTemplate exporter = new ExcelExporter();
    //导出路径
    var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExportByTemplate_Test) + ".xlsx");
    if (File.Exists(filePath)) File.Delete(filePath);

    //根据模板导出
    await exporter.ExportByTemplate(filePath, jobj, tplPath);

```
上述代码大家可以在单元测试`DynamicExportWithJObjectByTemplate_Test`中找到。

**值得注意的是，由于此处使用了`JObject`对象，因此在使用时需要按装包`Newtonsoft.Json`。但是，`Magicodes.IE.Excel`本身并不依赖`Newtonsoft.Json`。**

运行后可以看到如下图所示的结果：

![动态导出结果](../res/image-20210308180430331.png)

## 4.使用Dictionary<string, object>完成动态导出

导出的代码和上面是一样的，只是数据结构使用了`Dictionary`：

```csharp
var data = new Dictionary<string, object>()
{
    { "Company","雪雁" },
    { "Address", "湖南长沙" },
    { "Contact", "雪雁" },
    { "Tel", "136xxx" },
    { "BookInfos",new List<Dictionary<string,object>>()
        {
            new Dictionary<string, object>()
            {
                {"No","A1" },
                {"RowNo",1 },
                {"Name","Docker+Kubernetes应用开发与快速上云" },
                {"EditorInChief","李文强" },
                {"PublishingHouse","机械工业出版社" },
                {"Price",65 },
                {"PurchaseQuantity",50000 },
                {"Cover","https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965" },
                {"Remark","买起" }
            },
            new Dictionary<string, object>()
            {
                {"No","A2" },
                {"RowNo",2 },
                {"Name","Docker+Kubernetes应用开发与快速上云" },
                {"EditorInChief","李文强" },
                {"PublishingHouse","机械工业出版社" },
                {"Price",65 },
                {"PurchaseQuantity",50000 },
                {"Cover","https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965" },
                {"Remark","k8s真香" }
            }
        }
    }
};
//模板路径
var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
    "DynamicExportTpl.xlsx");
//创建Excel导出对象
IExportFileByTemplate exporter = new ExcelExporter();
//导出路径
var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExportWithDictionaryByTemplate_Test) + ".xlsx");
if (File.Exists(filePath)) File.Delete(filePath);

//根据模板导出
await exporter.ExportByTemplate(filePath, data, tplPath);
```

具体代码见`DynamicExportWithDictionaryByTemplate_Test`。

Tips:

- [如何使用集合初始值设定项初始化字典（C# 编程指南）](https://docs.microsoft.com/zh-cn/dotnet/csharp/programming-guide/classes-and-structs/how-to-initialize-a-dictionary-with-a-collection-initializer?WT.mc_id=DT-MVP-5004079)

## 5.使用ExpandoObject完成动态导出

同上，代码如下所示：

```csharp
dynamic data = new ExpandoObject();
data.Company = "雪雁";
data.Address = "湖南长沙";
data.Contact = "雪雁";
data.Tel = "136xxx";
data.BookInfos = new List<ExpandoObject>() { };

dynamic book1 = new ExpandoObject();
book1.No = "A1";
book1.RowNo = 1;
book1.Name = "Docker+Kubernetes应用开发与快速上云";
book1.EditorInChief = "李文强";
book1.PublishingHouse = "机械工业出版社";
book1.Price = 65;
book1.PurchaseQuantity = 50000;
book1.Cover = "https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965";
book1.Remark = "买买买";
data.BookInfos.Add(book1);

dynamic book2 = new ExpandoObject();
book2.No = "A2";
book2.RowNo = 2;
book2.Name = "Docker+Kubernetes应用开发与快速上云";
book2.EditorInChief = "李文强";
book2.PublishingHouse = "机械工业出版社";
book2.Price = 65;
book2.PurchaseQuantity = 50000;
book2.Cover = "https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965";
book2.Remark = "买买买";
data.BookInfos.Add(book2);

//模板路径
var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
    "DynamicExportTpl.xlsx");
//创建Excel导出对象
IExportFileByTemplate exporter = new ExcelExporter();
//导出路径
var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExportWithExpandoObjectByTemplate_Test) + ".xlsx");
if (File.Exists(filePath)) File.Delete(filePath);

//根据模板导出
await exporter.ExportByTemplate(filePath, data, tplPath);
```

具体代码参考`DynamicExportWithExpandoObjectByTemplate_Test`。

## 最后

本教程至此就结束了，如有疑问，麻烦大家多多提交Issue。

**Magicodes.IE：导入导出通用库，支持Dto导入导出、模板导出、花式导出以及动态导出，支持Excel、Csv、Word、Pdf和Html。**

- Github：<https://github.com/dotnetcore/Magicodes.IE>
- 码云（手动同步，不维护）：<https://gitee.com/magicodes/Magicodes.IE>

**相关库会一直更新，在功能体验上有可能会和本文教程有细微的出入，请以相关具体代码、版本日志、单元测试示例为准。**