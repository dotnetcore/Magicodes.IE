<img align="right" src="./res/logo.jpg" width="300"/>

# Magicodes.IE | [简体中文](README.md)
[![Member project of .NET Core Community](https://img.shields.io/badge/member%20project%20of-NCC-9e20c9.svg)](https://github.com/dotnetcore)
[![nuget](https://img.shields.io/nuget/v/Magicodes.IE.Core.svg?style=flat-square)](https://www.nuget.org/packages/Magicodes.IE.Core) 
[![Build Status](https://dev.azure.com/xinlaiopencode/Magicodes.IE/_apis/build/status/dotnetcore.Magicodes.IE?branchName=master)](https://dev.azure.com/xinlaiopencode/Magicodes.IE/_build/latest?definitionId=4&branchName=master)
[![stats](https://img.shields.io/nuget/dt/Magicodes.IE.Core.svg?style=flat-square)](https://www.nuget.org/stats/packages/Magicodes.IE.Core?groupby=Version)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/dotnetcore/Magicodes.IE/blob/master/LICENSE)

[![Stargazers over time](https://starchart.cc/dotnetcore/Magicodes.IE.svg)](https://starchart.cc/dotnetcore/Magicodes.IE)

![Azure DevOps tests (master)](https://img.shields.io/azure-devops/tests/xinlaiopencode/Magicodes.IE/4/master)
![Azure DevOps coverage (branch)](https://img.shields.io/azure-devops/coverage/xinlaiopencode/Magicodes.IE/4/master)
[![Financial Contributors on Open Collective](https://opencollective.com/magicodes/all/badge.svg?label=financial+contributors)](https://opencollective.com/magicodes) 

## Overview

Import and export general library, support Dto import and export, template export, fancy export and dynamic export, support Excel, Csv, Word, Pdf and Html.

**![General description](./docs/Magicodes.IE.en.png)**

## Milestone

|  #   |    Status     | Completion time | Milestone situation |
| :--: | :-----------: | :------: | :----------------------------------------------------------: |
| [2.5](https://github.com/dotnetcore/Magicodes.IE/issues?q=+is%3Aissue+milestone%3A2.5) | ☕In progress |2020-10-30| [待办](https://github.com/dotnetcore/Magicodes.IE/milestone/7) |
| [2.4](https://github.com/dotnetcore/Magicodes.IE/issues?q=+is%3Aissue+milestone%3A2.4) | 🚩Completed |2020-09-30| [已完成](https://github.com/dotnetcore/Magicodes.IE/milestone/6) |
| [2.3](https://github.com/dotnetcore/Magicodes.IE/issues?q=+is%3Aissue+milestone%3A2.3) | 🚩Completed |2020-06-30| [已完成](https://github.com/dotnetcore/Magicodes.IE/milestone/5) |
| [2.2](https://github.com/dotnetcore/Magicodes.IE/issues?q=+is%3Aissue+milestone%3A2.2) | 🚩Completed |2020-04-31| [已完成](https://github.com/dotnetcore/Magicodes.IE/milestone/4) |
| [2.1](https://github.com/dotnetcore/Magicodes.IE/issues?q=+is%3Aissue+milestone%3A2.1) | 🚩Completed |2020-03-15| [已完成](https://github.com/dotnetcore/Magicodes.IE/milestone/2?closed=1) |

### Azure DevOps
- Build Status：[![Build Status](https://dev.azure.com/xinlaiopencode/Magicodes.IE/_apis/build/status/dotnetcore.Magicodes.IE?branchName=master)](https://dev.azure.com/xinlaiopencode/Magicodes.IE/_build/latest?definitionId=4&branchName=master)
- Azure DevOps coverage (master):  ![Azure DevOps coverage (branch)](https://img.shields.io/azure-devops/coverage/xinlaiopencode/Magicodes.IE/4/master)
- Azure DevOps coverage (develop):  ![Azure DevOps coverage (branch)](https://img.shields.io/azure-devops/coverage/xinlaiopencode/Magicodes.IE/4/develop)
- Azure DevOps tests (master):  ![Azure DevOps tests (master)](https://img.shields.io/azure-devops/tests/xinlaiopencode/Magicodes.IE/4/master)
- Azure DevOps tests (develop):  ![Azure DevOps tests (develop)](https://img.shields.io/azure-devops/tests/xinlaiopencode/Magicodes.IE/4/develop)

For details, see: <https://dev.azure.com/xinlaiopencode/Magicodes.IE/_build?definitionId=4&_a=summary>

### Nuget

#### Stable version (recommended)

| **Name** | **Nuget** |
|----------|:-------------:|
| **Magicodes.IE.Core** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Core)](https://www.nuget.org/packages/Magicodes.IE.Core)** |
| **Magicodes.IE.Excel** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Excel)](https://www.nuget.org/packages/Magicodes.IE.Excel)**   |
| **Magicodes.IE.Pdf** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Pdf)](https://www.nuget.org/packages/Magicodes.IE.Pdf)**   |
| **Magicodes.IE.Word** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Word)](https://www.nuget.org/packages/Magicodes.IE.Word)**   |
| **Magicodes.IE.Html** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Html)](https://www.nuget.org/packages/Magicodes.IE.Html)**   |
| **Magicodes.IE.Csv** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Csv)](https://www.nuget.org/packages/Magicodes.IE.Csv)**   |
| **Magicodes.IE.AspNetCore** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.AspNetCore)](https://www.nuget.org/packages/Magicodes.IE.AspNetCore)**   |

### **Note**

- Excel import does not support ".xls" files, that is, Excel97-2003 is not supported. 
- For use in Docker, please refer to the section "Use in Docker" in the documentation. 
- Relevant functions have been compiled with unit tests. You can refer to unit tests during the use process. 

### **Tutorial**

Sorry, due to limited energy, please help translate.

1. **[基础教程之导入学生数据](docs/1.基础教程之导入学生数据.md "1.基础教程之导入学生数据") **

2. **[基础教程之导出Excel](docs/2.基础教程之导出Excel.md "2.基础教程之导出Excel")  [（点此访问国内文档）](https://docs.xin-lai.com/2020/02/19/%E7%BB%84%E4%BB%B6/Magicodes.IE/2.Magicodes.IE%E5%9F%BA%E7%A1%80%E6%95%99%E7%A8%8B%E4%B9%8B%E5%AF%BC%E5%87%BAExcel/)**

3. **[基础教程之导出Pdf收据](docs/3.基础教程之导出Pdf收据.md "3.基础教程之导出Pdf收据")** [**(点此访问国内文档)**](https://docs.xin-lai.com/2020/02/25/%E7%BB%84%E4%BB%B6/Magicodes.IE/3.Magicodes.IE%E5%9F%BA%E7%A1%80%E6%95%99%E7%A8%8B%E4%B9%8B%E5%AF%BC%E5%87%BAPdf/)

4. **[在Docker中使用](docs/4.在Docker中使用.md "4.在Docker中使用")**

5. **[动态导出](docs/5.动态导出.md "5.动态导出")**

6. **[多Sheet导入](docs/6.多Sheet导入.md "6.多Sheet导入")**
7. **[Csv导入导出](docs/7.Csv导入导出.md "7.Csv导入导出")**
8. **[Excel图片导入导出](docs/8.Excel图片导入导出.md "8.Excel图片导入导出")** [**(点此访问国内文档)**](https://docs.xin-lai.com/2020/03/16/%E6%8A%80%E6%9C%AF%E6%96%87%E6%A1%A3/%E4%BD%BF%E7%94%A8Magicodes.IE.Excel%E5%AE%8C%E6%88%90Excel%E5%9B%BE%E7%89%87%E7%9A%84%E5%AF%BC%E5%85%A5%E5%92%8C%E5%AF%BC%E5%87%BA/)

9. **[Excel模板导出之导出教材订购表](docs/9.Excel模板导出之导出教材订购表.md "9.Excel模板导出之导出教材订购表")（[点此访问国内文档](https://docs.xin-lai.com/2020/01/08/%E7%BB%84%E4%BB%B6/Magicodes.IE/7.Excel%E6%A8%A1%E6%9D%BF%E5%AF%BC%E5%87%BA%E4%B9%8B%E5%AF%BC%E5%87%BA%E6%95%99%E6%9D%90%E8%AE%A2%E8%B4%AD%E8%A1%A8/)）**

10. **[进阶篇之导入导出筛选器](https://docs.xin-lai.com/2020/09/21/%E7%BB%84%E4%BB%B6/Magicodes.IE/Magicodes.IE%E4%B9%8B%E5%AF%BC%E5%85%A5%E5%AF%BC%E5%87%BA%E7%AD%9B%E9%80%89%E5%99%A8/)**
11. **[Magicodes.IE之花式导出](https://docs.xin-lai.com/2020/09/28/%E7%BB%84%E4%BB%B6/Magicodes.IE/Magicodes.IE%E4%B9%8B%E8%8A%B1%E5%BC%8F%E5%AF%BC%E5%87%BA/)**
12. **[.NETCore中通过请求头导出多种格式文件](docs/12.NETCore中通过请求头导出多种格式文件.md)**
13. **[性能测试](docs/13.性能测试.md)**

**See below for other tutorials or unit tests**

**See below for update history. **

## Features

- **Need to be used in conjunction with related import and export DTO models, support import and export through DTO and related characteristics. Configure features to control related logic and display results without modifying the logic code;**
**![](./res/导入Dto.png "导入Dto")**
- **Support various filters to support scenarios such as multi-language, dynamic control column display, etc. For specific usage, see unit test:**
  - **Import column header filter (you can dynamically specify the imported column and imported value mapping relationship)**
  - **Export column header filter (can dynamically control the export column, support dynamic export (DataTable))**
  - **Import result filter (can modify annotation file)**
- **Export supports text custom filtering or processing;**
- **Import supports automatic skipping of blank lines in the middle;**
- **Import supports automatically generate import templates based on DTO, and automatically mark required items;**
![](./res/自动生成的导入模板.png "自动生成的导入模板")
- **Import supports data drop-down selection, currently only supports enumerated types;**
- **Imported data supports the processing of leading and trailing spaces and intermediate spaces, allowing specific columns to be set;**
- **Import supports automatic template checking, automatic data verification, unified exception handling, and unified error encapsulation, including exceptions, template errors and row data errors;**
![](./res/数据错误统一返回.png "数据错误")
- **Support import header position setting, the default is 1;**
- **Support import columns out of order, no need to correspond one to one in order;**
- **Support to import the specified column index, automatic recognition by default;**
- **Exporting Excel supports splitting of Sheets, only need to set the value of [MaxRowNumberOnASheet] of the characteristic [ExporterAttribute]. If it is 0, no splitting is required. See unit test for details;**
- **Support importing into Excel for error marking;**
![](./res/数据错误.png "数据错误标注")
![](./res/多个错误.png "多个错误")
- **Import supports cutoff column setting, if not set, blank cutoff will be encountered by default;**
- **Support exporting HTML, Word, Pdf, support custom export template;**
  -**Export HTML**
![](./res/导出html.png "导出HTML")
  -**Export Word**
![](./res/导出Word.png "导出Word")
  -**Export Pdf, support settings, see the update log for details**
![](./res/导出Pdf.png "导出Pdf")
  -**Export receipt**
![](./res/导出收据.png "导出收据.png")
- **Import supports repeated verification;**
![](./res/重复错误.png "重复错误.png")
- **Support single data template export, often used to export receipts, credentials and other businesses**
- **Support dynamic column export (based on DataTable), and the Sheet will be split automatically if it exceeds 100W. (Thanks to teacher Zhang Shanyou ([https://github.com/xin-lai/Magicodes.IE/pull/8](https://github.com/xin-lai/Magicodes.IE/pull/8) ))* *
- **Support dynamic/ExpandoObject dynamic column export**
```csharp
        [Fact(DisplayName = "DTO导出支持动态类型")]
        public async Task ExportAsByteArraySupportDynamicType_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(ExportAsByteArraySupportDynamicType_Test)}.xlsx");

            DeleteFile(filePath);

            var source = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();
            string fields = "text,number,name";
            var shapedData = source.ShapeData(fields) as ICollection<ExpandoObject>;

            var result = await exporter.ExportAsByteArray<ExpandoObject>(shapedData);
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            File.WriteAllBytes(filePath, result);
            File.Exists(filePath).ShouldBeTrue();
        }
```
- **Support value mapping, support setting value mapping relationship through "ValueMappingAttribute" feature. It is used to generate data validation constraints for import templates and perform data conversion. **
```csharp
        /// <summary>
        ///     性别
        /// </summary>
        [ImporterHeader(Name = "性别")]
        [Required(ErrorMessage = "性别不能为空")]
        [ValueMapping(text: "男", 0)]
        [ValueMapping(text: "女", 1)]
        public Genders Gender { get; set; }
```

- **Support the generation of imported data verification items of enumeration and Bool type, and related data conversion**
	- **Enumeration will automatically obtain the description, display name, name and value of the enumeration by default to generate data items**

		```csharp
			/// <summary>
			/// 学生状态 正常、流失、休学、勤工俭学、顶岗实习、毕业、参军
			/// </summary>
			public enum StudentStatus
			{
				/// <summary>
				/// 正常
				/// </summary>
				[Display(Name = "正常")]
				Normal = 0,

				/// <summary>
				/// 流失
				/// </summary>
				[Description("流水")]
				PupilsAway = 1,

				/// <summary>
				/// 休学
				/// </summary>
				[Display(Name = "休学")]
				Suspension = 2,

				/// <summary>
				/// 勤工俭学
				/// </summary>
				[Display(Name = "勤工俭学")]
				WorkStudy = 3,

				/// <summary>
				/// 顶岗实习
				/// </summary>
				[Display(Name = "顶岗实习")]
				PostPractice = 4,

				/// <summary>
				/// 毕业
				/// </summary>
				[Display(Name = "毕业")]
				Graduation = 5,

				/// <summary>
				/// 参军
				/// </summary>
				[Display(Name = "参军")]
				JoinTheArmy = 6,
			}
		```

		![](./res/enum.png "枚举转数据映射序列")

	- **The bool type will generate "yes" and "no" data items by default**
	- **If custom value mapping has been set, no default options will be generated**

- **Support excel multi-sheet import**
  **![](./res/multipleSheet.png "枚举转数据映射序列")**

- **Support Excel template export, and support image rendering**
  **![](./res/ExcelTplExport.png "Excel模板导出")**

  The rendering syntax is as follows:

  ```
    {{Company}}  //Cell rendering
    {{Table>>BookInfos|RowNo}} //Table rendering start syntax
    {{Remark|>>Table}}//Table rendering end syntax
    {{Image::ImageUrl?Width=50&Height=120&Alt=404}} //Picture rendering
    {{Image::ImageUrl?w=50&h=120&Alt=404}} //Picture rendering
    {{Image::ImageUrl?Alt=404}} //Picture rendering
  ```

  Custom pipelines will be supported in the future.

- **Support Excel import template to generate annotation**
  ![](./res/ImportLabel.png "Excel导入标注")

- **Support Excel image import and export**
  - Picture import
    - Import as Base64
    - Import to temporary directory
    - Import to the specified directory
   - Picture export
    - Export file path as picture
    - Export network path as picture

- **Support multiple entities to export multiple Sheets**

- **Support using some features under the System.ComponentModel.DataAnnotations namespace to control import and export** [#63](https://github.com/dotnetcore/Magicodes.IE/issues/63)

- **Support the use of custom formatter in ASP.NET Core Web API to export content such as Excel, Pdf, Csv** [#64](https://github.com/dotnetcore/Magicodes.IE/issues/64 )

### FAQ

[Question List](https://github.com/dotnetcore/Magicodes.IE/issues?q=label%3Aquestion)

### **Update history**

**[Update history](RELEASE.md)**

## Contributors

### Code Contributors

This project exists thanks to all the people who contribute. [[Contribute](CONTRIBUTING.md)].
<a href="https://github.com/dotnetcore/Magicodes.IE/graphs/contributors"><img src="https://opencollective.com/magicodes/contributors.svg?width=890&button=false" /></a>

### Financial Contributors

Become a financial contributor and help us sustain our community. [[Contribute](https://opencollective.com/magicodes/contribute)]

#### Individuals

<a href="https://opencollective.com/magicodes"><img src="https://opencollective.com/magicodes/individuals.svg?width=890"></a>

#### Organizations

Support this project with your organization. Your logo will show up here with a link to your website. [[Contribute](https://opencollective.com/magicodes/contribute)]

<a href="https://opencollective.com/magicodes/organization/0/website"><img src="https://opencollective.com/magicodes/organization/0/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/1/website"><img src="https://opencollective.com/magicodes/organization/1/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/2/website"><img src="https://opencollective.com/magicodes/organization/2/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/3/website"><img src="https://opencollective.com/magicodes/organization/3/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/4/website"><img src="https://opencollective.com/magicodes/organization/4/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/5/website"><img src="https://opencollective.com/magicodes/organization/5/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/6/website"><img src="https://opencollective.com/magicodes/organization/6/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/7/website"><img src="https://opencollective.com/magicodes/organization/7/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/8/website"><img src="https://opencollective.com/magicodes/organization/8/avatar.svg"></a>
<a href="https://opencollective.com/magicodes/organization/9/website"><img src="https://opencollective.com/magicodes/organization/9/avatar.svg"></a>
