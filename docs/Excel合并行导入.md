# Excel合并行导入

## 说明

Magicodes.IE.Excel目前已支持合并行单元格导入，如本篇教程所示。

## 安装包Magicodes.IE.Excel

```powershell
Install-Package Magicodes.IE.Excel
```

## 添加Dto

参考示例代码如下所示：

```csharp
public class MergeRowsImportDto
{
    [ImporterHeader(Name = "学号")]
    public long No { get; set; }

    [ImporterHeader(Name = "姓名")]
    public string Name { get; set; }

    [ImporterHeader(Name = "性别")]
    public string Sex { get; set; }
}
```
## 准备Excel导入文件

参考如图：

![导入文件](../res/image-20210306105147319.png)

该文件可以在测试工程中找到。

## 编写导入实现

导入代码和正常的导入没什么区别：

```csharp
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "合并行.xlsx");
        var import = await Importer.Import<MergeRowsImportDto>(filePath);
```
Tips：

- [Directory.GetCurrentDirectory 方法：获取应用程序的当前工作目录](https://docs.microsoft.com/zh-cn/dotnet/api/system.io.directory.getcurrentdirectory?view=net-5.0&WT.mc_id=DT-MVP-5004079)
- [Path.Combine 方法：将多个字符串组合成一个路径](https://docs.microsoft.com/zh-cn/dotnet/api/system.io.path.combine?view=net-5.0&WT.mc_id=DT-MVP-5004079)
- [IExcelImporter](https://github.com/dotnetcore/Magicodes.IE/blob/master/src/Magicodes.ExporterAndImporter.Excel/IExcelImporter.cs)
- [IImporter](https://github.com/dotnetcore/Magicodes.IE/blob/master/src/Magicodes.ExporterAndImporter.Core/IImporter.cs#L23:22)

上述代码大家可以在单元测试`MergeRowsImportTest`中找到。调试运行后可以看到如下图所示：

![合并行导入](../res/image-20210307180551091.png)

## 最后

本教程至此就结束了，如有疑问，麻烦大家多多提交Issue。

**Magicodes.IE：导入导出通用库，支持Dto导入导出、模板导出、花式导出以及动态导出，支持Excel、Csv、Word、Pdf和Html。**

- Github：<https://github.com/dotnetcore/Magicodes.IE>
- 码云（手动同步，不维护）：<https://gitee.com/magicodes/Magicodes.IE>

