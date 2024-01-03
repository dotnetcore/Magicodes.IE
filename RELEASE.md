# Release Log

## 2.7.5.0
**2024.01.03**

- 发布Magicodes.IE.Stash，具体见<https://github.com/dotnetcore/Magicodes.IE/pull/498>及<https://github.com/dotnetcore/Magicodes.IE/tree/master/src/Stash/Magicodes.IE.Stash>
- 增加导入Stream时的回调委托参数<https://github.com/dotnetcore/Magicodes.IE/commit/e9a2a1b3b493fea90379b0ee519b7bbb42011122>
- 增加ValueMappingsBase值映射特性基类，方便从外部(如：Abp框架)继承实现枚举、bool类型的多语言显示<https://github.com/dotnetcore/Magicodes.IE/pull/544>

## 2.7.4.4
**2023.04.22**

- 修复获取导入错误文件流时，因批注行导致的无表头的问题

## 2.7.4.3
**2023.03.16**

- 优化列名中不包含FieldErrors的key触发的异常信息
- Excel模板导出，图片格式支持base64

## 2.7.4.2
**2023.01.04**

- 重构表头排序

## 2.7.4.1
**2023.01.02**

- Excel导出去除Task.FromResult包装，增加异步操作
- 修改内部集合类，充分利用池化内存，降低内存分配

## 2.7.4
**2022.12.30**

- Fix DateTime cell convert to Nullable<DateTime>，见pr [#475](https://github.com/dotnetcore/Magicodes.IE/pull/475)。

## 2.7.3
**2022.12.27**

- 更新Haukcode.WkHtmlToPdfDotNet到1.5.86
- tagetFrameworks移除net461

## 2.7.2
**2022.12.04**

- 修复FontSize的Bug

## 2.7.1
**2022.12.01**

- Magicodes.IE.EPPlus默认添加SkiaSharp.NativeAssets.Linux.NoDependencies包，以便于在Linux环境下使用
- 导入验证支持将错误数据通过Stream的方式返回，感谢sampsonye （见pr[#466](https://github.com/dotnetcore/Magicodes.IE/pull/466)）

## 2.7.0
**2022.11.07**

- 添加SkiaSharp
- 移除SixLabors.Fonts
- 感谢linch90的大力支持（具体见pr[#462](https://github.com/dotnetcore/Magicodes.IE/pull/462)） 
- 部分方法改为虚方法

## 2.7.0-beta
**2022.10.27**

-  使用SixLabors.ImageSharp替代System.Drawing，感谢linch90 （见pr[#454](https://github.com/dotnetcore/Magicodes.IE/pull/454)）

## 2.6.9
**2022.10.26**

-  fix: 动态数据源导出到多个sheet的问题 （见[#449](https://github.com/dotnetcore/Magicodes.IE/issues/449)）

## 2.6.8
**2022.10.18**

-  Excel模板导出添加API，以支持通过文件流模板：Task<byte[]> ExportBytesByTemplate<T>(T data, Stream templateStream)

## 2.6.7
**2022.10.12**

- ExporterHeaderFilter支持修改列索引，以支持动态排序，需设置ExporterHeaderAttribute.ColumnIndex属性（注意不应修改Index属性），值范围为0~10000。设置错误会自动调整到相近的边界值。
- 提供ExporterHeadersFilter筛选器，以支持批量修改列头。
- 重构、优化列排序代码。

## 2.6.5-beta1
**2022.07.17**

- 【修复】如果为动态类型导出，如datatable/dynamic/proxy等，会将原始数据转成字符串。
- fix: 修复没有正确释放Graphics对象的问题 （见PR[#401](https://github.com/dotnetcore/Magicodes.IE/pull/401)）
- feat(module: excel): Export of the byte type Enum value is allowed （见PR[#367](https://github.com/dotnetcore/Magicodes.IE/pull/367)）
- feat(module: excel): The export can be of Nullable Enum type （见PR[#398](https://github.com/dotnetcore/Magicodes.IE/pull/398)）
- fix(module: Excel): Excel ParseData

## 2.6.4
**2022.04.17**

- 优化了ColumnIndex在生成模板时的实现，增加了ColumnIndex的单测（见PR[#385](https://github.com/dotnetcore/Magicodes.IE/pull/385)）。

- 添加了NPOI的独立扩展包——Magicodes.IE.Excel.NPOI，以便于后续给用户提供更多的支持。目前仅提供了 SaveToExcelWithXSSFWorkbook 扩展方法。

- 修复RequiredIfAttribute的Bug。

- 修复导出JPG图片在Linux环境下可能引起的无限循环的问题（见PR[#396](https://github.com/dotnetcore/Magicodes.IE/pull/396)）。

- Excel图片导入时，图片列支持为空。

- 更新CsvHelper到最新版本，并修改相关代码。

## 2.6.3
**2022.03.06**

- 完善筛选器注册机制，在指定了特性ImportHeaderFilter、ExporterHeaderFilter等值后，筛选器将匹配对于的类型（见PR[#384](https://github.com/dotnetcore/Magicodes.IE/pull/384)），如不指定则作为全局筛选器。如下述代码，注入了多个同类型的筛选器，通过指定了ImportHeaderFilter限制了此Dto仅使用ImportHeaderFilterB：
```csharp
    builder.Services.AddTransient<IImportHeaderFilter, ImportHeaderFilterA>();
    builder.Services.AddTransient<IImportHeaderFilter, ImportHeaderFilterB>();
    builder.Services.AddTransient<IImportHeaderFilter, ImportHeaderFilterC>();

    [ExcelImporter(ImportHeaderFilter = typeof(ImportHeaderFilterB))]
    public class ImportExcelTemplateDto
    {
        [ImporterHeader(Name = "TypeName")]
        public string? Name { get; set; }
    }
```
    


## 2.6.2
**2022.03.02**

- Excel导入时增加回调函数，方便增加自定义验证（见PR[#369](https://github.com/dotnetcore/Magicodes.IE/pull/369)）：

```csharp
        [Fact(DisplayName = "导入结果回调函数测试")]
        public async Task ImportResultCallBack_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "缴费流水导入模板.xlsx");
            var import = await Importer.Import<ImportPaymentLogDto>(filePath, (importResult) =>
            {
                int rowNum = 2;//首行数据对应Excel中的行号
                foreach (var importPaymentLogDto in importResult.Data)
                {
                    if (importPaymentLogDto.Amount > 5000)
                    {
                        var dataRowError = new DataRowErrorInfo();
                        dataRowError.RowIndex = rowNum;
                        dataRowError.FieldErrors.Add("Amount", "金额不能大于5000");
                        importResult.RowErrors.Add(dataRowError);
                    }
                    rowNum++;
                }
                return importResult;
            });
            import.ShouldNotBeNull();
            import.HasError.ShouldBeTrue();
            import.RowErrors.ShouldContain(p => p.RowIndex == 3 && p.FieldErrors.ContainsKey("金额不能大于5000"));
            import.Exception.ShouldBeNull();
            import.Data.Count.ShouldBe(20);
        }
```

- 优化获取DisplayName的逻辑（见PR[#372](https://github.com/dotnetcore/Magicodes.IE/pull/372)）

- 导出CSV支持ColumnIndex（见PR[#381](https://github.com/dotnetcore/Magicodes.IE/pull/381)）

- 优化Pdf导出逻辑，统一各平台导出代码

## 2.6.1

- 修复内存未及时回收

## 2.6.0
**2021.11.28**

- 添加两个动态验证特性（见PR[#319 by Afonsof91](https://github.com/dotnetcore/Magicodes.IE/pull/359)）：

  - 添加特性`DynamicStringLengthAttribute`,以便支持动态配置字符串长度验证。使用参考：

  ```csharp
  public class DynamicStringLengthImportDto
  {
      [ImporterHeader(Name = "名称")]
      [Required(ErrorMessage = "名称不能为空")]
      [DynamicStringLength(typeof(DynamicStringLengthImportDtoConsts), nameof(DynamicStringLengthImportDtoConsts.MaxNameLength), ErrorMessage = "名称字数不能超过{1}")]
      public string Name { get; set; }
  }
  
  public static class DynamicStringLengthImportDtoConsts
  {
      public static int MaxNameLength { get; set; } = 3;
  }
  ```

  - 添加特性`RequiredIfAttribute`，以支持动态开启必填验证。使用参考：

  ```csharp
  public class RequiredIfAttributeImportDto
  {
      [ImporterHeader(Name = "名称是否必填")]
      [Required(ErrorMessage = "名称是否必填不能为空")]
      [ValueMapping("是", true)]
      [ValueMapping("否", false)]
      public bool IsNameRequired { get; set; }
  
      [ImporterHeader(Name = "名称")]
      [RequiredIf("IsNameRequired", "True", ErrorMessage = "名称不能为空")]
      [MaxLength(10, ErrorMessage = "名称字数超出最大值：10")]
      public string Name { get; set; }
  }
  ```

- CSV添加对分隔符的配置，具体见PR[#319 by Afonsof91](https://github.com/dotnetcore/Magicodes.IE/pull/319)

- Excel导入添加对`TimeSpan`类型的支持，使用参考`TimeSpan_Test`

- 初步添加对.NET6的适配

## 2.5.6.3
**2021.10.23**

- 导出日期格式化支持`DateTimeOffset`类型，具体见PR[#349](https://github.com/dotnetcore/Magicodes.IE/pull/349)，感谢[YaChengMu](https://github.com/YaChengMu)
- 修改Magicodes.IE.EPPlus的包依赖PR[#351](https://github.com/dotnetcore/Magicodes.IE/pull/351)

## 2.5.6.2
**2021.10.13**
- 支持自定义列字体颜色，具体见PR[#342](https://github.com/dotnetcore/Magicodes.IE/pull/342)，感谢[xiangxiren](https://github.com/xiangxiren)
- 修复日期格式化的问题，具体见PR[#344](https://github.com/dotnetcore/Magicodes.IE/pull/344)，感谢[ccccccmd](https://github.com/ccccccmd)

## 2.5.6.1
**2021.10.06**
- 修复 [#337](https://github.com/dotnetcore/Magicodes.IE/issues/377)，bool?类型导出的映射问题

## 2.5.6.0
**2021.10.05**
- 合并Magicodes.EPPlus到Magicodes.IE，修复所有单元测试并修复部分Bug

## 2.5.5.4
**2021.09.02**
- 修复可为空枚举导入时的验证问题[#322](https://github.com/dotnetcore/Magicodes.IE/issues/322)。

## 2.5.5.3
**2021.08.27**
- 修复Append方式导出多个sheet时，发生“Tablename is not unique”错误，具体见[#299](https://github.com/dotnetcore/Magicodes.IE/issues/299)。

## 2.5.5.2
**2021.08.24**
- 添加对Abp模块的包装，具体见[#318](https://github.com/dotnetcore/Magicodes.IE/issues/318)342。

## 2.5.5.1
**2021.08.07**
- 为了简化ASP.NET Core下的Excel导出，对Excel导出进行了进一步的封装
- 添加`Magicodes.IE.Excel.AspNetCore`工程，添加`XlsxFileResult`的Action Result，支持泛型集合、Bytes数组、Steam直接导出
- 修改部分命名和命名空间

## 2.5.4.9

**2021.07.23**
- 修复Excel合并行导入在存在空的合并单元格时可能的数据读取错误[#305](https://github.com/dotnetcore/Magicodes.IE/issues/305)



## 2.5.4.8

**2021.07.15**
- Magicodes.EPPlus回退到4.6.6，以修复格式错乱的问题
- 修复Excel仅导出错误数据时的Bug[#302](https://github.com/dotnetcore/Magicodes.IE/pull/302)
- 完善多语言[#298](https://github.com/dotnetcore/Magicodes.IE/pull/298)，以及完善单元测试

## 2.5.4.6

**2021.07.04**
- 模板导出支持一行多个表格[#296](https://github.com/dotnetcore/Magicodes.IE/issues/296)

## 2.5.4.5

**2021.06.29**
- 合并PR[#295](https://github.com/dotnetcore/Magicodes.IE/pull/295),完善模板导出类型定义的问题

## 2.5.4.4

**2021.06.25**
- Fix only first [ColumnIndex] is valid exception[#289](https://github.com/dotnetcore/Magicodes.IE/issues/289)

## 2.5.4.3

**2021.06.18**
- Update ImportTestColumnIndex_Test
- Magicodes.EPPlus was upgraded to 4.6.7[#285](https://github.com/dotnetcore/Magicodes.IE/issues/285)

## 2.5.4.2

**2021.06.05**
- Fix ImporterHeader->ColumnIndex
- Utilize RecyclableMemoryStream instead of "new MemoryStream" all over[#282](https://github.com/dotnetcore/Magicodes.IE/issues/282)

## 2.5.4.1

**2021.06.05**
- EXCEL模板导出支持XOffset和YOffset[#280](https://github.com/dotnetcore/Magicodes.IE/issues/280)
- EXCEL修复ValueMapping
- Core工程多语言配置
- EXCEL优化时间导出

## 2.5.4.0

**2021.06.01**

- EXCEL支持自动换行属性[#278](https://github.com/dotnetcore/Magicodes.IE/issues/278)
- EXCEL支持隐藏列属性[#273](https://github.com/dotnetcore/Magicodes.IE/issues/273)
- EXCEL优化时间优化

## 2.5.3.9

**2021.05.26**
- 修复ValueMappingAttribute[#272](https://github.com/dotnetcore/Magicodes.IE/issues/272)

## 2.5.3.8

**2021.05.10**

- Excel模板导出功能，将单行复制改为多行复制
- PDF导出内存优化

## 2.5.3.7

**2021.04.23**
- 修复导入模板生成，格式错误[#261](https://github.com/dotnetcore/Magicodes.IE/issues/261)
例如：

## 2.5.3.6

**2021.04.18**
- 支持对导入模板生成，预设值单元格格式[#253](https://github.com/dotnetcore/Magicodes.IE/issues/253)
例如：
```
[ImporterHeader(Name = "序号", Format ="@")]
```
- 单元格图片导出支持偏移设置[#250](https://github.com/dotnetcore/Magicodes.IE/issues/250)
例如：
```
**YOffset**：垂直偏移（可进行移动图片）
**XOffset**：水平偏移（可进行移动图片）
```
- 支持多sheet导入SheetIndex的支持[#254](https://github.com/dotnetcore/Magicodes.IE/issues/254)
例如：
```
[ExcelImporter(SheetIndex = 2)]
```

## 2.5.3.5

**2021.04.13**
- Excel导入支持列头忽略大小写导入（全局配置：IsIgnoreColumnCase）

## 2.5.3.4

**2021.04.06**
- Excel导入修复枚举值不在范围时的错误提示

## 2.5.3.3

**2021.04.03**
- Excel导入逻辑移除5万行的限制，默认不限制导入数量


## 2.5.3.2

**2021.03.30**
- Excel修复OutputBussinessErrorData扩展方法
- 多Sheet导入对Stream的支持


## 2.5.3.1

**2021.03.12**

- Excel模板导出支持使用Dictionary、ExpandoObject完成动态导出
- 优化模板导出逻辑

## 2.5.3

**2021.03.08**

- Excel模板导出支持使用JSON对象完成动态导出 [#I398DI](https://gitee.com/magicodes/Magicodes.IE/issues/I398DI)

## 2.5.2

**2021.03.05**

- Excel导入支持合并行数据 [#239](https://github.com/dotnetcore/Magicodes.IE/issues/239)

## 2.5.1.8
 **2021.02.23**
- Input string was not in a correct format.[#241](https://github.com/dotnetcore/Magicodes.IE/issues/241)
- 使用Stream方式导入xlsx，rowErrors里的rowIndex位置不对[#236](https://github.com/dotnetcore/Magicodes.IE/issues/236)

## 2.5.1.7
**2021.02.20**
- Excel支持Base64导出 [#219](https://github.com/dotnetcore/Magicodes.IE/issues/219)
- 修复 [#214](https://github.com/dotnetcore/Magicodes.IE/issues/214)

## 2.5.1.6

**2021.01.31**
- 部分重构模板导出
- Excel模板导出语法解析加强 [#211](https://github.com/dotnetcore/Magicodes.IE/issues/211)
- 修复当表格下面存在变量时，无法渲染的Bug

## 2.5.1.5

**2021.01.29**
- 移除模板导出时的控制台日志输出

## 2.5.1.4

**2021.01.09**
- 修复Excel导出列头索引与内容排序不一致问题及单测  [#226](https://github.com/dotnetcore/Magicodes.IE/issues/226)

## 2.5.1.3

**2021.01.02**
- Add PDF support for paper size
- Add PDF support for margins [#223](https://github.com/dotnetcore/Magicodes.IE/issues/223)

## 2.5.1

**2020.12.21**
- 导出支持使用ColumnIndex指定导出顺序，以导出时在某些情况下顺序不一致的问题（Export supports the use of ColumnIndex to specify the export order, so that the order is inconsistent in some cases when exporting）  [#179](https://github.com/dotnetcore/Magicodes.IE/issues/179)

## 2.5.0

**2020.12.03**

- Excel导出支持HeaderRowIndex [#164](https://github.com/dotnetcore/Magicodes.IE/issues/164)
- 增加Excel枚举导出对DescriptionAttribute的支持 [#168](https://github.com/dotnetcore/Magicodes.IE/issues/168)
- Excel生成导入模板支持内置数据验证[#167](https://github.com/dotnetcore/Magicodes.IE/issues/167)
  - 支持数据验证
    - 支持MaxLengthAttribute、MinLengthAttribute、StringLengthAttribute、RangeAttribute
  - 支持输入提示 
  To fix The Mapping Values of The total length of a Data Validation list always exceed 255 characters (# 196) (https://github.com/dotnetcore/Magicodes.IE/issues/196)
- Excel export List data type errors, and formatting issues.[#191](https://github.com/dotnetcore/Magicodes.IE/issues/191) [193] (https://github.com/dotnetcore/Magicodes.IE/issues/193)
- 导入Excel对Enum类型匹配值映射时，忽略值前后空格
- fix MappingValues The total length of a DataValidation list cannot exceed 255 characters [#196](https://github.com/dotnetcore/Magicodes.IE/issues/196)
- Excel导出List数据类型存在错误，以及格式化问题。 [#191](https://github.com/dotnetcore/Magicodes.IE/issues/191) [#193](https://github.com/dotnetcore/Magicodes.IE/issues/193)
- The ColumnIndex property does not appear to be valid in Excel import  [#198](https://github.com/dotnetcore/Magicodes.IE/issues/198)
- TableStyle修改为枚举类型

## 2.5.0-beta6
**2020.11.26**
 - The ColumnIndex property does not appear to be valid in Excel import  [#198](https://github.com/dotnetcore/Magicodes.IE/issues/198)

## 2.5.0-beta5
**2020.11.25**
- fix MappingValues The total length of a DataValidation list cannot exceed 255 characters [#196](https://github.com/dotnetcore/Magicodes.IE/issues/196)
- Excel导出List数据类型存在错误，以及格式化问题。 [#191](https://github.com/dotnetcore/Magicodes.IE/issues/191) [#193](https://github.com/dotnetcore/Magicodes.IE/issues/193)

## 2.5.0-beta4
**2020.11.20**
To fix The Mapping Values of The total length of a Data Validation list always exceed 255 characters (# 196) (https://github.com/dotnetcore/Magicodes.IE/issues/196)
- Excel export List data type errors, and formatting issues. [#191](https://github.com/dotnetcore/Magicodes.IE/issues/191) [193] (https://github.com/dotnetcore/Magicodes.IE/issues/193)
- 导入Excel对Enum类型匹配值映射时，忽略值前后空格

## 2.5.0-beta3
**2020.10.29**
- Excel生成导入模板支持内置数据验证[#167](https://github.com/dotnetcore/Magicodes.IE/issues/167)
  - 支持数据验证
    - 支持MaxLengthAttribute、MinLengthAttribute、StringLengthAttribute、RangeAttribute
  - 支持输入提示 


## 2.5.0-beta2
**2020.10.20**
- Excel导出支持HeaderRowIndex [#164](https://github.com/dotnetcore/Magicodes.IE/issues/164)
- 增加Excel枚举导出对DescriptionAttribute的支持 [#168](https://github.com/dotnetcore/Magicodes.IE/issues/168)

## 2.4.0
**2020.10.01**
- 支持单元格导出宽度设置 [#129](https://github.com/dotnetcore/Magicodes.IE/issues/129)
- Excel导出支持对Enum的ValueMapping设置 [#106](https://github.com/dotnetcore/Magicodes.IE/issues/106)
- Excel导出支持对bool类型的ValueMapping设置 [#16](https://github.com/dotnetcore/Magicodes.IE/issues/16)
- [#152](https://github.com/dotnetcore/Magicodes.IE/issues/152) 筛选器支持依赖注入
 ```csharp
    public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
    {
        AppDependencyResolver.Init(app.ApplicationServices);
        //all other code
    }
 ```
- #151 导出添加AutoFitMaxRows，超过指定行数则不启用AutoFit
- 添加全局IsDisableAllFilter属性，以通过特性禁用所有筛选器
- [#142](https://github.com/dotnetcore/Magicodes.IE/issues/142) 【修复】根据模板列表高度的设置，统一设置渲染高度
- [#157](https://github.com/dotnetcore/Magicodes.IE/issues/157)【修复】对低版本框架的兼容
- Excel导入对图片获取算法的优化

## 2.4.0-beta4
**2020.09.26**
- [#157](https://github.com/dotnetcore/Magicodes.IE/issues/157)【修复】对低版本框架的兼容

## 2.4.0-beta3
**2020.09.24**
- [#142](https://github.com/dotnetcore/Magicodes.IE/issues/142) 【修复】根据模板列表高度的设置，统一设置渲染高度

## 2.4.0-beta2
**2020.09.16**
- [#152](https://github.com/dotnetcore/Magicodes.IE/issues/152) 筛选器支持依赖注入
 ```csharp
    public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
    {
        AppDependencyResolver.Init(app.ApplicationServices);
        //all other code
    }
 ```
- #151 导出添加AutoFitMaxRows，超过指定行数则不启用AutoFit
- 添加全局IsDisableAllFilter属性，以通过特性禁用所有筛选器

## 2.4.0-beta1
**2020.09.14**
- 支持单元格导出宽度设置 [#129](https://github.com/dotnetcore/Magicodes.IE/issues/129)
- Excel导出支持对Enum的ValueMapping设置 [#106](https://github.com/dotnetcore/Magicodes.IE/issues/106)
- Excel导出支持对bool类型的ValueMapping设置 [#16](https://github.com/dotnetcore/Magicodes.IE/issues/16)

## 2.3.0
**2020.08.30**

## 2.3.0-beta8
**2020.08.22**
- 修复基于文件流导入时的NULL异常，并完善单元测试 [#141](https://github.com/dotnetcore/Magicodes.IE/issues/141)**

## 2.3.0-beta7
**2020.08.16**
- excel添加对ExpandoObject类型的支持 [#135](https://github.com/dotnetcore/Magicodes.IE/issues/135)**

**2020.08.10**
- **【Nuget】版本更新到2.3.0-beta6**
- 多Sheet导入保存标注错误单元测试，并没出现多数据导入效验bug [#108](https://github.com/dotnetcore/Magicodes.IE/issues/108)
- Excel多Sheet 导入模板生成 [#133](https://github.com/dotnetcore/Magicodes.IE/issues/133)
- 修复Excel模板图片高度问题 [#131](https://github.com/dotnetcore/Magicodes.IE/issues/131)

**2020.08.04**
- **【Nuget】版本更新到2.3.0-beta5**
- **在runtimes native包问题**
- **对于跨平台native中 `COM Interop is not supported on this platform.`修复** [#130](https://github.com/dotnetcore/Magicodes.IE/issues/130)

**2020.07.14**
- **【Nuget】版本更新到2.3.0-beta4**

**2020.07.13**
- **【Nuget】版本更新到2.3.0-beta3**
- **【PDF导出】修复Linux下导出PDf 出错问题 [#125](https://github.com/dotnetcore/Magicodes.IE/issues/125)**

**2020.07.06**

- **【Nuget】版本更新到2.3.0-beta2**
- **【Excel导出】导出业务错误数据支持直接返回错误数据的文件流字节**
- **【Excel导出】对追加sheet实现同一个Model可自定义传入不同sheet名称**

     -  exporter.Append(list1,"sheet1").SeparateBySheet().Append(list2).ExportAppendData(filePath);
     
- **【Nuget】针对于一些客户端不支持SemVer 2.0.0 进行采取兼容机制**

**2020.06.22**

- **【Nuget】版本更新到2.3.0-beta1**
- **【Excel导出】添加对Excel模板导出函数的支持**

          - {{Formula::AVERAGE?params=G4:G6}}
          - {{Formula::SUM?params=G4:G6&G4}}

**2020.06.16**

- **【Nuget】版本更新到2.2.6**
- **【HTML导出】添加对NETCore2.2模板引擎的支持**

**2020.06.14**

- **【Nuget】版本更新到2.2.5**
- **【Excel导出】增加分栏、分sheet、追加rows导出 [#74](https://github.com/dotnetcore/Magicodes.IE/issues/74)**

      - exporter.Append(list1).SeparateByColumn().Append(list2).ExportAppendData(filePath);
      - exporter.Append(list1).SeparateBySheet().Append(list2).ExportAppendData(filePath);
      - exporter.Append(list1).SeparateByRow().AppendHeaders().Append(list2).ExportAppendData(filePath);
- **[Excel导出】修复‘IsAllowRepeat＝true’ [#107](https://github.com/dotnetcore/Magicodes.IE/issues/107)**
- **[Pdf导出】增加PDF扩展方法，支持通过以参数形式传递特性参数 [#104](https://github.com/dotnetcore/Magicodes.IE/issues/104)**

      - Task<byte[]> ExportListBytesByTemplate<T>(ICollection＜T＞ data, PdfExporterAttribute pdfExporterAttribute,string temple);
      - Task<byte[]> ExportBytesByTemplate＜T＞(T data, PdfExporterAttribute pdfExporterAttribute,string template);

**2020.06.07**

- **【Nuget】版本更新到2.2.4**
- **【Excel导入】增加`导入失败`仅返回错误行功能**
- **【Excel导入】修复导入的空行标注位置偏移**
- **【Excel导出】增加`SeparateByColumn`进行分割追加列**  

#### 2020.05.31

- **【Nuget】版本更新到2.2.3**
- **【Excel导入】增加了stream Csv导入扩展方法**
- **【Word导出】修复word文件字节导出错误**


#### **2020.05.24**

- **【Nuget】版本更新到2.2.2**
- **【Excel导入】增加了stream导入扩展方法**
- **【Excel导出】增加了内容居中（单列居中、整表居中）**
- **【导出】对一些中间件代码进行了修复及优化**

#### **2020.05.16**
- **【Nuget】版本更新到2.2.1**
- **【PDF导出】对模板引擎进行升级更新**


#### **2020.05.12**

- **【Nuget】版本更新到2.2.0**
- **【Excel模板导出】支持导出字节**
- **【文档】Magicodes.IE Csv导入导出**
- **【Excel导入导出】修复标注的添加问题**
- **【导出】ASP.NET Core Web API 中使用自定义格式化程序导出Excel、Pdf、Csv等内容** [#64](https://github.com/dotnetcore/Magicodes.IE/issues/64)
- **【导入导出】支持使用System.ComponentModel.DataAnnotations命名空间下的部分特性来控制导入导出**  [#63](https://github.com/dotnetcore/Magicodes.IE/issues/63)

#### **2020.04.16**
- **【Nuget】版本更新到2.2.0-beta9**
- **【Excel模板导出】修复只存在一列时的导出 [#73](https://github.com/dotnetcore/Magicodes.IE/issues/73)**
- **【Excel导入】支持返回表头和索引 [#76](https://github.com/dotnetcore/Magicodes.IE/issues/76)**
- **【Excel导入导入】[#63](https://github.com/dotnetcore/Magicodes.IE/issues/63)**
  - 支持使用System.ComponentModel.DataAnnotations命名空间下的部分特性来控制导入导出，比如
    - DisplayAttribute
    - DisplayFormatAttribute
    - DescriptionAttribute
  - 封装简单的易于使用的单一特性，例如
    - IEIgnoreAttribute（可作用于属性、枚举成员，可影响导入和导出）

#### **2020.04.02**
- **【Nuget】版本更新到2.2.0-beta8**

- **【Excel模板导出】支持图片 [#62](https://github.com/dotnetcore/Magicodes.IE/issues/62)，渲染语法如下所示：**

 ```
  {{Image::ImageUrl?Width=50&Height=120&Alt=404}}
  {{Image::ImageUrl?w=50&h=120&Alt=404}}
  {{Image::ImageUrl?Alt=404}}
 ```

#### **2020.03.29**
- **【Nuget】版本更新到2.2.0-beta7**
- **【Excel模板导出】修复渲染问题 [#51](https://github.com/dotnetcore/Magicodes.IE/issues/51)**

#### **2020.03.27**
- **【Nuget】版本更新到2.2.0-beta6**
- **【Excel导入导出】修复.NET Core 2.2的包引用问题 [#68](https://github.com/dotnetcore/Magicodes.IE/issues/68)**

#### **2020.03.26**
- **【Nuget】版本更新到2.2.0-beta4**
- **【Excel多Sheet导出】修复[#66](https://github.com/dotnetcore/Magicodes.IE/issues/66)，并添加单元测试**

#### **2020.03.25**
- **【Nuget】版本更新到2.2.0-beta3**
- **【Excel导入】修复日期问题 [#68](https://github.com/dotnetcore/Magicodes.IE/issues/68)**
- **【Excel导出】添加ExcelOutputType设置，支持输出无格式的导出。[#54](https://github.com/dotnetcore/Magicodes.IE/issues/54)可以使用此方式。**

#### **2020.03.19**
- **【Nuget】版本更新到2.2.0-beta2**
- **【Excel导入】修复日期格式的导入Bug，支持DateTime和DateTimeOffset以及可为空类型，默认支持本地化时间格式（默认根据地区自动使用本地日期时间格式）**
- **【Excel导入导出】添加单元测试ExportAndImportUseOneDto_Test，对使用同一个Dto导出并导入进行测试。Issue见 [#53](https://github.com/dotnetcore/Magicodes.IE/issues/53)**

#### **2020.03.18**
- **【Nuget】版本更新到2.2.0-beta1**
- **【Excel导出】添加以下API:**
````csharp

        /// <summary>
        ///     追加集合到当前导出程序
        ///     append the collection to context
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        ExcelExporter Append<T>(ICollection<T> dataItems) where T : class;

        /// <summary>
        ///     导出所有的追加数据
        ///     export excel after append all collectioins
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportAppendData(string fileName);

        /// <summary>
        ///     导出所有的追加数据
        ///     export excel after append all collectioins
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        Task<byte[]> ExportAppendDataAsByteArray();

````

- **【Excel导出】支持多个实体导出多个Sheet**，感谢@ccccccmd 的贡献 [#pr52](https://github.com/dotnetcore/Magicodes.IE/pull/52) ，Issue见 [#50](https://github.com/dotnetcore/Magicodes.IE/issues/50)。使用代码参考，具体见单元测试（ExportMutiCollection_Test）：

````csharp
            var exporter = new ExcelExporter();
            var list1 = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>();
            var list2 = GenFu.GenFu.ListOf<ExportTestDataWithSplitSheet>(30);
            var result = exporter.Append(list1).Append(list2).ExportAppendData(filePath);
````

#### **2020.03.12**
- **【Nuget】版本更新到2.1.4**
- **【Excel导入】支持图片导入，见特性ImportImageFieldAttribute**
  - 导入为Base64
  - 导入到临时目录
  - 导入到指定目录
- **【Excel导出】支持图片导出，见特性ExportImageFieldAttribute**
  - 将文件路径导出为图片
  - 将网络路径导出为图片

#### **2020.03.06**
- **【Nuget】版本更新到2.1.3**
- **【Excel导入】修复GUID类型的问题。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/44>）。**

#### **2020.02.25**
- **【Nuget】版本更新到2.1.2**
- **【导入导出】已支持CSV**
- **【文档】完善Pdf导出文档**

#### **2020.02.24**
- **【Nuget】版本更新到2.1.1-beta**
- **【导入】Excel导入支持导入标注，仅需设置ExcelImporterAttribute的ImportDescription属性，即会在顶部生成Excel导入说明**
- **【重构】添加两个接口**
  - IExcelExporter：继承自IExporter, IExportFileByTemplate，Excel特有的API将在此补充
  - IExcelImporter：继承自IImporter，Excel特有的API在此补充，例如“ImportMultipleSheet”、“ImportSameSheets”
- **【重构】增加实例依赖注入**
- **【构建】完成代码覆盖率的DevOps的配置**

#### **2020.02.14**
- **【Nuget】版本更新到2.1.0**
- **【导出】PDF导出支持.NET 4.6.1，具体见单元测试**

#### **2020.02.13**
- **【Nuget】版本更新到2.0.2**
- **【导入】修复单列导入的Bug，单元测试“OneColumnImporter_Test”。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/35>）。**
- **【导出】修复导出HTML、Pdf、Word时，模板在某些情况下编译报错的问题。**
- **【导入】重写空行检查。**

#### **2020.02.11**
- **【Nuget】版本更新到2.0.0**
- **【导出】Excel模板导出修复多个Table渲染以及合并单元格渲染的问题，具体见单元测试“ExportByTemplate_Test1”。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/34>）。**
- **【导出】完善模板导出的单元测试，针对导出结果添加渲染检查，确保所有单元格均已渲染。**

#### **2020.02.05**
- **【Nuget】版本更新到2.0.0-beta4**
- **【导入】支持列筛选器（需实现接口【IImportHeaderFilter】），可用于兼容多语言导入等场景，具体见单元测试【ImportHeaderFilter_Test】**
- **【导入】支持传入标注文件路径，不传参则默认同目录"_"后缀保存**
- **【导入】完善单元测试【ImportResultFilter_Test】**
- **【其他】修改【ValueMappingAttribute】的命名空间为Magicodes.ExporterAndImporter.Core**

#### **2020.02.04**
- **【Nuget】版本更新到2.0.0-beta2**
- **【导入】支持导入结果筛选器——IImportResultFilter，可用于多语言场景的错误标注，具体使用见单元测试【ImportResultFilter_Test】**
- **【其他】修改IExporterHeaderFilter的命名空间为Magicodes.ExporterAndImporter.Core.Filters**

#### **2020.01.18**
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

#### **2020.01.16**
- **【Nuget】版本更新到1.4.25**
- **【导出】修复没有定义导出特性会报错的情形，具体见单元测试“ExportTestDataWithoutExcelExporter_Test”。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/21>）。**

#### **2020.01.16**
- **【Nuget】版本更新到1.4.24**
- **【导出】修复日期格式默认导出数字的Bug，默认输出“yyyy-MM-dd”，可以通过设置“[ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]”来修改。问题见（<https://github.com/dotnetcore/Magicodes.IE/issues/22>）。**

#### **2020.01.14**
- **【Nuget】版本更新到1.4.21**
- **【导出】Excel模板导出修复数据项为Null报错的Bug。**

#### **2020.01.09**
- **【Nuget】版本更新到1.4.20**
- **【导出】Excel模板导出性能优化。5000条表格数据1秒内完成，具体见单元测试ExportByTemplate_Large_Test。**

#### **2020.01.08**
- **【Nuget】版本更新到1.4.18**
- **【导入】支持导入最大数量限制**
    - **ImporterAttribute支持MaxCount设置，默认为50000**
    - **完成相关单元测试**

#### **2020.01.07**
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
