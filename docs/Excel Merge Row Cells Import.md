# Excel Merge Row Cells Import

## Description

This tutorial introduces Magicodes.IE.Excel already supports merged row cell import.

## Installation package Magicodes.IE.Excel

```powershell
Install-Package Magicodes.IE.Excel
```

## Add Dto

The reference sample code is shown below.

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
## Prepare Excel import file

As shown below:

![导入文件](../res/image-20210306105147319.png)

This file can be found in the test project.

## Write import implementation

Importing the code is no different from other normal import codes.

```csharp
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "合并行.xlsx");
        var import = await Importer.Import<MergeRowsImportDto>(filePath);
```
The above code can be found in the unit test `MergeRowsImportTest`. After debugging and running, you can see the following figure.

![合并行导入](../res/image-20210307180551091.png)

## Finally

This tutorial ends here, if you have questions, please submit more Issue.

**Magicodes.IE：导入导出通用库，支持Dto导入导出、模板导出、花式导出以及动态导出，支持Excel、Csv、Word、Pdf和Html。**

- Github：<https://github.com/dotnetcore/Magicodes.IE>
- gitee(Manual sync, no maintenance)：<https://gitee.com/magicodes/Magicodes.IE>

