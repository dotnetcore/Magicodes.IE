# Magicodes.IE

导入导出通用库，支持Dto导入导出以及动态导出，支持Excel、Word、Pdf和Html。

[Magicodes.IE 2.0发布](https://docs.xin-lai.com/2020/02/12/%E7%BB%84%E4%BB%B6/Magicodes.IE/Magicodes.IE%202.0%E5%8F%91%E5%B8%83/)

- Github：<https://github.com/dotnetcore/Magicodes.IE>
- 码云（手动同步，不维护）：<https://gitee.com/magicodes/Magicodes.IE>

## 疯狂的徽章

### GitHub

- ![GitHub contributors](https://img.shields.io/github/contributors/dotnetcore/Magicodes.IE?style=social) ![GitHub license badge](https://img.shields.io/github/license/dotnetcore/Magicodes.IE?style=social) ![GitHub repo size](https://img.shields.io/github/repo-size/dotnetcore/Magicodes.IE?style=social)
- ![GitHub commit activity](https://img.shields.io/github/commit-activity/m/dotnetcore/Magicodes.IE?style=social)  ​![GitHub last commit](https://img.shields.io/github/last-commit/dotnetcore/Magicodes.IE?style=social)
- ![GitHub issues badge](https://img.shields.io/github/issues/dotnetcore/Magicodes.IE?style=social) ![GitHub issues badge](https://img.shields.io/github/issues-closed/dotnetcore/Magicodes.IE?style=social)
- ![GitHub forks badge](https://img.shields.io/github/forks/dotnetcore/Magicodes.IE?style=social)	![GitHub stars](https://img.shields.io/github/stars/dotnetcore/Magicodes.IE?style=social)
- ![GitHub pull requests](https://img.shields.io/github/issues-pr/dotnetcore/Magicodes.IE?style=social)	![GitHub closed pull requests](https://img.shields.io/github/issues-pr-closed/dotnetcore/Magicodes.IE?style=social)

### Azure DevOps 
- Build Status：[![Build Status](https://dev.azure.com/xinlaiopencode/Magicodes.IE/_apis/build/status/dotnetcore.Magicodes.IE?branchName=master)](https://dev.azure.com/xinlaiopencode/Magicodes.IE/_build/latest?definitionId=4&branchName=master)
- Azure DevOps coverage (master):  ![Azure DevOps coverage (branch)](https://img.shields.io/azure-devops/coverage/xinlaiopencode/Magicodes.IE/4/master) 
- Azure DevOps tests (master):  ![Azure DevOps tests (master)](https://img.shields.io/azure-devops/tests/xinlaiopencode/Magicodes.IE/4/master)
- Azure DevOps tests (develop):  ![Azure DevOps tests (develop)](https://img.shields.io/azure-devops/tests/xinlaiopencode/Magicodes.IE/4/develop)

具体见：<https://dev.azure.com/xinlaiopencode/Magicodes.IE/_build?definitionId=4&_a=summary>

### Nuget

| **名称** |      **Nuget**      |
|----------|:-------------:|
| **Magicodes.IE.Core** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Core)](https://www.nuget.org/packages/Magicodes.IE.Core)** |
| **Magicodes.IE.Excel** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Excel)](https://www.nuget.org/packages/Magicodes.IE.Excel)**   |
| **Magicodes.IE.Pdf** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Pdf)](https://www.nuget.org/packages/Magicodes.IE.Pdf)**   |
| **Magicodes.IE.Word** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Word)](https://www.nuget.org/packages/Magicodes.IE.Word)**   |
| **Magicodes.IE.Html** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Html)](https://www.nuget.org/packages/Magicodes.IE.Html)**   |
| **Magicodes.IE.Csv** | **[![NuGet](https://buildstats.info/nuget/Magicodes.IE.Csv)](https://www.nuget.org/packages/Magicodes.IE.Csv)**   |

### **注意**

- **Excel导入不支持“.xls”文件，即不支持Excel97-2003。**
- **如需在Docker中使用，请参阅文档中的《Docker中使用》一节。**
- **相关功能均已编写单元测试，在使用的过程中可以参考单元测试。**
- **此库会长期支持，但是由于精力有限，希望大家能够多多参与。**

### **教程**

1. **[基础教程之导入学生数据](docs/1.基础教程之导入学生数据.md "1.基础教程之导入学生数据")  （[点此访问国内文档](https://docs.xin-lai.com/2019/11/26/%E7%BB%84%E4%BB%B6/Magicodes.IE/1.%E5%9F%BA%E7%A1%80%E6%95%99%E7%A8%8B%E4%B9%8B%E5%AF%BC%E5%85%A5%E5%AD%A6%E7%94%9F%E6%95%B0%E6%8D%AE/)）**

2. **[基础教程之导出Excel](docs/2.基础教程之导出Excel.md "2.基础教程之导出Excel")  [（点此访问国内文档）](https://docs.xin-lai.com/2020/02/19/%E7%BB%84%E4%BB%B6/Magicodes.IE/2.Magicodes.IE%E5%9F%BA%E7%A1%80%E6%95%99%E7%A8%8B%E4%B9%8B%E5%AF%BC%E5%87%BAExcel/)**

3. **[基础教程之导出Pdf收据](docs/3.基础教程之导出Pdf收据.md "3.基础教程之导出Pdf收据")** [**(点此访问国内文档)**](https://docs.xin-lai.com/2020/02/25/%E7%BB%84%E4%BB%B6/Magicodes.IE/3.Magicodes.IE%E5%9F%BA%E7%A1%80%E6%95%99%E7%A8%8B%E4%B9%8B%E5%AF%BC%E5%87%BAPdf/)

4. **[在Docker中使用](docs/4.在Docker中使用.md "4.在Docker中使用")**

5. **动态导出（待补充）**

6. **多Sheet导入（待补充）**
7. **Csv导入导出（待补充）**

8. **[Excel模板导出之导出教材订购表](docs/7.Excel模板导出之导出教材订购表.md "7.Excel模板导出之导出教材订购表")（[点此访问国内文档](https://docs.xin-lai.com/2020/01/08/%E7%BB%84%E4%BB%B6/Magicodes.IE/7.Excel%E6%A8%A1%E6%9D%BF%E5%AF%BC%E5%87%BA%E4%B9%8B%E5%AF%BC%E5%87%BA%E6%95%99%E6%9D%90%E8%AE%A2%E8%B4%AD%E8%A1%A8/)）**

9. **进阶篇之导入导出筛选器（待补充）**
10. **主体API说明**
11. **其他教程见下文或单元测试**

**更新历史见下文。**



### **特点**

**![总体说明](./docs/Magicodes.IE.png)**

- **需配合相关导入导出的DTO模型使用，支持通过DTO以及相关特性控制导入导出。配置特性即可控制相关逻辑和显示结果，无需修改逻辑代码；**
**![](./res/导入Dto.png "导入Dto")**
- **支持各种筛选器，以便支持多语言、动态控制列展示等场景，具体使用见单元测试：**
  - **导入列头筛选器（可动态指定导入列、导入的值映射关系）**
  - **导出列头筛选器（可动态控制导出列，支持动态导出（DataTable））**
  - **导入结果筛选器（可修改标注文件）**
- **导出支持文本自定义过滤或处理；**
- **导入支持中间空行自动跳过；**
- **导入支持自动根据 DTO 生成导入模板,针对必填项将自动标注；**
**![](./res/自动生成的导入模板.png "自动生成的导入模板")**
- **导入支持数据下拉选择，目前仅支持枚举类型；**
- **导入数据支持前后空格以及中间空格处理，允许指定列进行设置；**
- **导入支持模板自动检查，数据自动校验，异常统一处理，并提供统一的错误封装，包含异常、模板错误和行数据错误；**
**![](./res/数据错误统一返回.png "数据错误")**
- **支持导入表头位置设置，默认为1；**
- **支持导入列乱序，无需按顺序一一对应；**
- **支持导入指定列索引，默认自动识别；**
- **导出Excel支持拆分Sheet，仅需设置特性【ExporterAttribute】的【MaxRowNumberOnASheet】的值，为0则不拆分。具体见单元测试；**
- **支持将导入Excel进行错误标注；**
**![](./res/数据错误.png "数据错误标注")**
**![](./res/多个错误.png "多个错误")**
- **导入支持截止列设置，如未设置则默认遇到空格截止；**
- **支持导出HTML、Word、Pdf，支持自定义导出模板；**
  - **导出HTML**
**![](./res/导出html.png "导出HTML")**
  - **导出Word**
**![](./res/导出Word.png "导出Word")**
  - **导出Pdf，支持设置，具体见更新日志**
**![](./res/导出Pdf.png "导出Pdf")**
  - **导出收据**
**![](./res/导出收据.png "导出收据.png")**
- **导入支持重复验证；**
**![](./res/重复错误.png "重复错误.png")**
- **支持单个数据模板导出，常用于导出收据、凭据等业务**
- **支持动态列导出（基于DataTable），并且超过100W将自动拆分Sheet。（感谢张善友老师（[https://github.com/xin-lai/Magicodes.IE/pull/8](https://github.com/xin-lai/Magicodes.IE/pull/8 ) ））**
- **支持值映射，支持通过“ValueMappingAttribute”特性设置值映射关系。用于生成导入模板的数据验证约束以及进行数据转换。**
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

- **支持枚举和Bool类型的导入数据验证项的生成，以及相关数据转换**
	- **枚举默认情况下会自动获取枚举的描述、显示名、名称和值生成数据项**

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

		**![](./res/enum.png "枚举转数据映射序列")**

	- **bool类型默认会生成“是”和“否”的数据项**
	- **如果已设置自定义值映射，则不会生成默认选项**
- **支持excel多Sheet导入；**
  **![](./res/multipleSheet.png "枚举转数据映射序列")**
- **支持Excel模板导出**
  **![](./res/ExcelTplExport.png "Excel模板导出")**


### **VNext**

> **以下内容均已有思路，但是缺乏精力，因此虚席待PR，有兴趣的朋友可以参与进来，多多交流。**

- [ ] **将代码单元测试覆盖率提高到90%（目前为86%）**
- [x] **Pdf导出支持.NET Framework 461**
- [x] **完成自动构建流程，并通过自动构建发包**
- [ ] **表头样式设置**
- [x] **自定义模板导出**
  - [x] **Excel （[#10](https://github.com/dotnetcore/Magicodes.IE/issues/10)）**
- [x] **加强值映射序列，比如支持方法、Dto接口的方式来获取**
- [ ] **生成导入模板时必填项支持自定义样式配置**
- [x] **CSV支持**
- [x] **Sheet拆分（有兴趣的朋友可以参考张队的PR：[https://github.com/xin-lai/Magicodes.IE/pull/14](https://github.com/xin-lai/Magicodes.IE/pull/14)）**
- [ ] **Excel导出支持图片**
- [x] **解决Excel导出无法进行数据筛选的问题（[#17](https://github.com/dotnetcore/Magicodes.IE/issues/17)）**
- [ ] **Excel单元格自动合并（[#9](https://github.com/dotnetcore/Magicodes.IE/issues/9)）**
- [ ] **导入导出支持指定位置[CellAddress(Row = 2, Column = 2)]（[#19](https://github.com/dotnetcore/Magicodes.IE/issues/19)）**
- [ ] **生成的导入模板支持数据验证**
- [ ] **优化包依赖，拆解项目**
- [x] **导入结果筛选器**
- [x] **导入列头筛选器**

### **联系我们**

> ##### **订阅号**

**关注“麦扣聊技术”订阅号可以获得最新文章、教程、文档：**

**![](./res/wechat.jpg "麦扣聊技术")**

> ##### **QQ群**

- **编程交流群<85318032>**

- **产品交流群<897857351>**

> ##### **文档官网&官方博客**

- **文档官网：<https://docs.xin-lai.com/>**
- **博客：<http://www.cnblogs.com/codelove/>**


> ##### **其他开源库**

- **<https://github.com/xin-lai>**
- **<https://gitee.com/magicodes>**


### **更新历史**

#### **2019.02.25**
- **【Nuget】版本更新到2.1.2**
- **【导入导出】已支持CSV**
- **【文档】完善Pdf导出文档**

#### **2019.02.24**
- **【Nuget】版本更新到2.1.1-beta**
- **【导入】Excel导入支持导入标注，仅需设置ExcelImporterAttribute的ImportDescription属性，即会在顶部生成Excel导入说明**
- **【重构】添加两个接口**
  - IExcelExporter：继承自IExporter, IExportFileByTemplate，Excel特有的API将在此补充
  - IExcelImporter：继承自IImporter，Excel特有的API在此补充，例如“ImportMultipleSheet”、“ImportSameSheets” 
- **【重构】增加实例依赖注入**
- **【构建】完成代码覆盖率的DevOps的配置**

#### **2019.02.14**
- **【Nuget】版本更新到2.1.0**
- **【导出】PDF导出支持.NET 4.6.1，具体见单元测试**

#### **2019.02.13**
- **【Nuget】版本更新到2.0.2**
- **【导入】修复单列导入的Bug，单元测试“OneColumnImporter_Test”。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/35>）。**
- **【导出】修复导出HTML、Pdf、Word时，模板在某些情况下编译报错的问题。**
- **【导入】重写空行检查。**

#### **2019.02.11**
- **【Nuget】版本更新到2.0.0**
- **【导出】Excel模板导出修复多个Table渲染以及合并单元格渲染的问题，具体见单元测试“ExportByTemplate_Test1”。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/34>）。**
- **【导出】完善模板导出的单元测试，针对导出结果添加渲染检查，确保所有单元格均已渲染。**

#### **2019.02.05**
- **【Nuget】版本更新到2.0.0-beta4**
- **【导入】支持列筛选器（需实现接口【IImportHeaderFilter】），可用于兼容多语言导入等场景，具体见单元测试【ImportHeaderFilter_Test】**
- **【导入】支持传入标注文件路径，不传参则默认同目录"_"后缀保存**
- **【导入】完善单元测试【ImportResultFilter_Test】**
- **【其他】修改【ValueMappingAttribute】的命名空间为Magicodes.ExporterAndImporter.Core**

#### **2019.02.04**
- **【Nuget】版本更新到2.0.0-beta2**
- **【导入】支持导入结果筛选器——IImportResultFilter，可用于多语言场景的错误标注，具体使用见单元测试【ImportResultFilter_Test】**
- **【其他】修改IExporterHeaderFilter的命名空间为Magicodes.ExporterAndImporter.Core.Filters**

#### **2019.01.18**
- **【Nuget】版本更新到2.0.0-beta1**
- **【导出】完全重构整个导出Excel模块并且重写大部分接口**
- **【导出】支持列头筛选器——IExporterHeaderFilter，具体使用见单元测试**
- **【导出】修复转换DataTable时支持为空类型**
- **【导出】导出Excel支持拆分Sheet，仅需设置特性【ExporterAttribute】的【MaxRowNumberOnASheet】的值，为0则不拆分。具体见单元测试**
- **【导出】修复导出结果无法筛选的问题。目前导出即为数据表**
- **【导出】添加扩展方法ToExcelExportFileInfo**
- **【导出】IExporter再添加两个动态DataTable导出方法，无需定义Dto即可动态导出数据，并且支持表头筛选器、Sheet拆分**
````csharp
        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="dataItems">数据</param>
        /// <param name="exporterHeaderFilter">表头筛选器</param>
        /// <param name="maxRowNumberOnASheet">一个Sheet最大允许的行数，设置了之后将输出多个Sheet</param>
        /// <returns>文件</returns>
        Task<ExportFileInfo> Export(string fileName, DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000);

        /// <summary>
        ///     导出Excel
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <param name="exporterHeaderFilter">表头筛选器</param>
        /// <param name="maxRowNumberOnASheet">一个Sheet最大允许的行数，设置了之后将输出多个Sheet</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportAsByteArray(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000);
````

#### **2019.01.16**
- **【Nuget】版本更新到1.4.25**
- **【导出】修复没有定义导出特性会报错的情形，具体见单元测试“ExportTestDataWithoutExcelExporter_Test”。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/21>）。**

#### **2019.01.16**
- **【Nuget】版本更新到1.4.24**
- **【导出】修复日期格式默认导出数字的Bug，默认输出“yyyy-MM-dd”，可以通过设置“[ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]”来修改。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/22>）。**

#### **2019.01.14**
- **【Nuget】版本更新到1.4.21**
- **【导出】Excel模板导出修复数据项为Null报错的Bug。**

#### **2019.01.09**
- **【Nuget】版本更新到1.4.20**
- **【导出】Excel模板导出性能优化。5000条表格数据1秒内完成，具体见单元测试ExportByTemplate_Large_Test。**

#### **2019.01.08**
- **【Nuget】版本更新到1.4.18**
- **【导入】支持导入最大数量限制**
    - **ImporterAttribute支持MaxCount设置，默认为50000**
    - **完成相关单元测试**

#### **2019.01.07**
- **【Nuget】版本更新到1.4.17**
- **【重构】重构IExportFileByTemplate中的ExportByTemplate，将参数htmlTemplate改为template。以便支持Excel模板导出。**
- **【导出】支持Excel模板导出并填写相关单元测试，如何使用见教程《Excel模板导出之导出教材订购表》**
    - **支持单元格单个绑定**
    - **支持列表**


#### **2019.12.17**
- **【Nuget】版本更新到1.4.16**
- **【导入】Excel导入支持多sheet导入，感谢tanyongzheng（[https://github.com/dotnetcore/Magicodes.IE/pull/18](https://github.com/dotnetcore/Magicodes.IE/pull/18)）**

#### **2019.12.10**
- **【Nuget】版本更新到1.4.15**
- **【测试】单元测试添加多框架版本支持 (<https://docs.xin-lai.com/2019/12/10/%E6%8A%80%E6%9C%AF%E6%96%87%E6%A1%A3/Magicodes.IE%E7%BC%96%E5%86%99%E5%A4%9A%E6%A1%86%E6%9E%B6%E7%89%88%E6%9C%AC%E6%94%AF%E6%8C%81%E5%92%8C%E6%89%A7%E8%A1%8C%E5%8D%95%E5%85%83%E6%B5%8B%E8%AF%95/>)**
- **【修复】修复部分.NET Framework 461下的问题**

#### **2019.12.06**
- **【Nuget】版本更新到1.4.14**
- **【重构】大量重构**
	- **移除部分未使用的代码**
	- **将TemplateFileInfo重命名为ExportFileInfo**
	- **将IExporterByTemplate接口拆分为4个接口：IExportListFileByTemplate, IExportListStringByTemplate, IExportStringByTemplate, IExportFileByTemplate，并修改相关实现**
	- **重构ImportHelper部分代码**
- **【导入】修复导入Excel时表头设置的问题，已对此编写单元测试，见【产品信息导入】**
- **【完善】编写ExportAsByteArray对于DataTable的单元测试，ExportWordFileByTemplate_Test**

#### **2019.11.25**
- **【Nuget】版本更新到1.4.13**
- **【导出】Pdf导出支持特性配置，详见单元测试【导出竖向排版收据】。目前主要支持以下设置：**
	- **Orientation：排版方向（横排、竖排）**
	- **PaperKind：纸张类型，默认A4**
	- **IsEnablePagesCount：是否启用分页数**
	- **Encoding：编码设置，默认UTF8**
	- **IsWriteHtml：是否输出HTML模板，如果启用，则会输出.html后缀的对应的HTML文件，方便调错**
	- **HeaderSettings：头部设置，通常可以设置头部的分页内容和信息**
	- **FooterSettings：底部设置**

#### **2019.11.24**
- **【Nuget】版本更新到1.4.12**
- **【导出】导出动态类支持超过100W数据时自动拆分Sheet（具体见PR：[https://github.com/xin-lai/Magicodes.IE/pull/14](https://github.com/xin-lai/Magicodes.IE/pull/14)）**

#### **2019.11.20**
- **【Nuget】版本更新到1.4.11**
- **【导出】修复Datatable列的顺序和DTO的顺序不一致，导致数据放错列（具体见PR：[https://github.com/xin-lai/Magicodes.IE/pull/13](https://github.com/xin-lai/Magicodes.IE/pull/13)）**

#### **2019.11.16**
- **【Nuget】版本更新到1.4.10**
- **【导出】修复Pdf导出在多线程下的问题**

#### **2019.11.13**
- **【Nuget】版本更新到1.4.5**
- **【导出】修复导出Pdf在某些情况下可能会导致内存报错的问题**
- **【导出】添加批量导出收据单元测试示例，并添加大量数据样本进行测试**

#### **2019.11.5**
- **【Nuget】版本更新到1.4.4**
- **【导入】修复枚举类型的问题，并编写单元测试**
- **【导入】增加值映射，支持通过“ValueMappingAttribute”特性设置值映射关系。用于生成导入模板的数据验证约束以及进行数据转换。**
- **【导入】优化枚举和Bool类型的导入数据验证项的生成，以便于模板生成和数据转换**
	- **枚举默认情况下会自动获取枚举的描述、显示名、名称和值生成数据项**
	- **bool类型默认会生成“是”和“否”的数据项**
	- **如果已设置自定义值映射，则不会生成默认选项**
- **【导入】支持枚举可为空类型**

#### **2019.10.30**
- **【Nuget】版本更新到1.4.0**
- **【导出】Excel导出支持动态列导出（基于DataTable），感谢张善友（https://github.com/xin-lai/Magicodes.IE/pull/8 ）**

#### **2019.10.22**
- **【Nuget】版本更新到1.3.7**
- **【导入】修复忽略列的验证问题**
- **【导入】修正验证错误信息，一行仅允许存在一条数据**
- **【导入】修复忽略列在某些情况下可能引发的异常**
- **【导入】添加存在忽略列的导入情形下的单元测试**

#### **2019.10.21**
- **【Nuget】版本更新到1.3.4**
- **【导入】支持设置忽略列，以便于在Dto定义数据列做处理或映射**

#### **2019.10.18**
- **【优化】优化.NET标准库2.1下集合转DataTable的性能**
- **【重构】多处IList<T>修改为ICollection<T>**
- **【完善】补充部分单元测试**

#### **2019.10.12**
- **【重构】重构HTML、PDF导出等逻辑，并修改IExporterByTemplate为：**
  - **Task<string> ExportListByTemplate<T>(IList<T> dataItems, string htmlTemplate = null) where T : class;**
  - **Task<string> ExportByTemplate<T>(T data, string htmlTemplate = null) where T : class;**
- **【示例】添加收据导出的单元测试示例**



#### **2019.9.28**
- **【导出】修改默认的导出HTML、Word、Pdf模板**
- **【导入】添加截断行的单元测试，以测试中间空格和结尾空格**
- **【导入】将【数据错误检测】和【导入】单元测试的Dto分开，确保全部单元测试通过**
- **【文档】更新文档**

#### **2019.9.26**
- **【导出】支持导出Word、Pdf、HTML，支持自定义导出模板**
- **【导出】添加相关导出的单元测试**
- **【导入】支持重复验证，需设置ImporterHeader特性的IsAllowRepeat为false**

#### **2019.9.19**
- **【导入】支持截止列设置，如未设置则默认遇到空格截止**
- **【导入】导入支持通过特性设置Sheet名称**

#### **2019.9.18**

- **【导入】重构导入模块**
- **【导入】统一导入错误消息**
	- **Exception ：导入异常信息**
	- **RowErrors ： 数据错误信息**
	- **TemplateErrors ：模板错误信息，支持错误分级**
	- **HasError : 是否存在错误（仅当出现异常并且错误等级为Error时返回true）**
- **【导入】基础类型必填自动识别，比如int、double等不可为空类型自动识别，无需额外设置Required**
- **【导入】修改Excel模板的Sheet名称**
- **【导入】支持导入表头位置设置，默认为1**
- **【导入】支持列乱序（导入模板的列序号不再需要固定）**
- **【导入】支持列索引设置**
- **【导入】支持将导入的Excel进行错误标注，支持多个错误**
- **【导入】加强对基础类型和可为空类型的支持**
- **【EPPlus】由于EPPlus.Core已经不维护，将EPPlus的包从EPPlus.Core改为EPPlus，**

#### **2019.9.11**

- **【导入】导入支持自动去除前后空格，默认启用，可以针对列进行关闭，具体见AutoTrim设置**
- **【导入】导入Dto的字段允许不设置ImporterHeader，支持通过DisplayAttribute特性获取列名**
- **【导入】导入的Excel移除对Sheet名称的约束，默认获取第一个Sheet**
- **【导入】导入增加对中间空格的处理支持，需设置FixAllSpace**
- **【导入】导入完善对日期类型的支持**
- **【导入】完善导入的单元测试**
