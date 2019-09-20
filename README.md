# Magicodes.ExporterAndImporter

导入导出通用库，目前仅支持导入导出Excel。

### 特点

- 封装导入导出业务,目前仅支持 Excel,后续将支持 CSV 以及其他业务
- 配置特性即可控制相关逻辑和显示结果，无需修改逻辑代码
- 推荐配合 DTO 使用，支持通过DTO以及相关特性控制导入导出
- 导出支持列头自定义处理以便支持多语言等场景
- 导出支持文本自定义过滤或处理
- 导入支持自动根据 DTO 生成导入模板,针对必填项将自动标注
![](./res/自动生成的导入模板.png "自动生成的导入模板")
- 导入支持数据下拉选择，目前仅支持枚举类型
- 导入数据支持前后空格以及中间空格处理，允许指定列进行设置
- 导入支持模板自动检查，数据自动校验，异常统一处理，并提供统一的错误封装，包含异常、模板错误和行数据错误
![](./res/数据错误统一返回.png "数据错误")
- 支持导入表头位置设置，默认为1
- 支持导入列乱序，无需按顺序一一对应
- 支持导入指定列索引，默认自动识别
- 支持将导入Excel进行错误标注
![](./res/数据错误.png "数据错误标注")
![](./res/多个错误.png "多个错误")
- 导入支持截止列设置，如未设置则默认遇到空格截止

### 相关官方Nuget包

| 名称     |      Nuget      |
|----------|:-------------:|
| Magicodes.IE.Core  |  [![NuGet](https://buildstats.info/nuget/Magicodes.IE.Core)](https://www.nuget.org/packages/Magicodes.IE.Core) |
| Magicodes.IE.Excel |    [![NuGet](https://buildstats.info/nuget/Magicodes.IE.Excel)](https://www.nuget.org/packages/Magicodes.IE.Excel)   |

### 联系我们

> ##### 订阅号

关注“麦扣聊技术”订阅号可以获得最新文章、教程、文档：
![](./res/wechat.jpg"麦扣聊技术")

> ##### QQ群

- 编程交流群<85318032>

- 产品交流群<897857351>

> ##### 官方博客

<http://www.cnblogs.com/codelove/>

> ##### 其他开源库

- <https://github.com/xin-lai>
- <https://gitee.com/magicodes>

### VNext

- 生成导入模板时必填项支持自定义配置
- 对模板错误进行标注
- CSV支持
- Word支持
- Pdf支持
- HTML
- 导入结果支持生成HTML输出
- 导入重复验证

### 更新历史

#### 2019.9.19
- 【导入】支持截止列设置，如未设置则默认遇到空格截止
- 【导入】导入支持通过特性设置Sheet名称

#### 2019.9.18

- 【导入】重构导入模块
- 【导入】统一导入错误消息
	- Exception ：导入异常信息
	- RowErrors ： 数据错误信息
	- TemplateErrors ：模板错误信息，支持错误分级
	- HasError : 是否存在错误（仅当出现异常并且错误等级为Error时返回true）
- 【导入】基础类型必填自动识别，比如int、double等不可为空类型自动识别，无需额外设置Required
- 【导入】修改Excel模板的Sheet名称
- 【导入】支持导入表头位置设置，默认为1
- 【导入】支持列乱序（导入模板的列序号不再需要固定）
- 【导入】支持列索引设置
- 【导入】支持将导入的Excel进行错误标注，支持多个错误
- 【导入】加强对基础类型和可为空类型的支持
- 【EPPlus】由于EPPlus.Core已经不维护，将EPPlus的包从EPPlus.Core改为EPPlus，

#### 2019.9.11

- 【导入】导入支持自动去除前后空格，默认启用，可以针对列进行关闭，具体见AutoTrim设置
- 【导入】导入Dto的字段允许不设置ImporterHeader，支持通过DisplayAttribute特性获取列名
- 【导入】导入的Excel移除对Sheet名称的约束，默认获取第一个Sheet
- 【导入】导入增加对中间空格的处理支持，需设置FixAllSpace
- 【导入】导入完善对日期类型的支持
- 【导入】完善导入的单元测试

### 导出 Demo

---
#### Demo1-1

普通导出
![](./res/1.png "Demo1-1")

>

    public class ExportTestData
    {
        public string Name1 { get; set; }
        public string Name2 { get; set; }
        public string Name3 { get; set; }
        public string Name4 { get; set; }
    }
    
    var result = await Exporter.Export(filePath, new List<ExportTestData>()
    {
        new ExportTestData()
        {
            Name1 = "1",
            Name2 = "test",
            Name3 = "12",
            Name4 = "11",
        },
        new ExportTestData()
        {
            Name1 = "1",
            Name2 = "test",
            Name3 = "12",
            Name4 = "11",
        }
    });

---
#### Demo1-2

特性导出
![](./res/2.png "Demo1-2")

>

    [ExcelExporter(Name = "测试", TableStyle = "Light10")]
    public class ExportTestDataWithAttrs
    {
        [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
        public string Text { get; set; }
    
        [ExporterHeader(DisplayName = "普通文本")]
        public string Text2 { get; set; }
    
        [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
        public string Text3 { get; set; }
    
        [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
        public double Number { get; set; }
    
        [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
        public string Name { get; set; }
    }
            var result = await Exporter.Export(filePath, new List<ExportTestDataWithAttrs>()
            {
                new ExportTestDataWithAttrs()
                {
                    Text = "啊实打实大苏打撒",
                    Name="aa",
                    Number =5000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
               new ExportTestDataWithAttrs()
                {
                    Text = "啊实打实大苏打撒",
                    Name="啊实打实大苏打撒",
                    Number =6000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
               new ExportTestDataWithAttrs()
                {
                    Text = "啊实打实速度大苏打撒",
                    Name="萨达萨达",
                    Number =6000,
                    Text2 = "突然他也让他人",
                    Text3 = "sadsad打发打发士大夫的"
                },
            });

#### Demo1-3

列头处理或者多语言支持
![](./res/3.png "Demo1-3")

>

    [ExcelExporter(Name = "测试", TableStyle = "Light10")]
    public class AttrsLocalizationTestData
    {
        [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
        public string Text { get; set; }
    
        [ExporterHeader(DisplayName = "普通文本")]
        public string Text2 { get; set; }
    
        [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
        public string Text3 { get; set; }
    
        [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
        public double Number { get; set; }
    
        [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
        public string Name { get; set; }
    }
            ExcelBuilder.Create().WithLocalStringFunc((key) =>
            {
                if (key.Contains("文本"))
                {
                    return "Text";
                }
                return "未知语言";
            }).Build();
    
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "testAttrsLocalization.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);
    
            var result = await Exporter.Export(filePath, new List<AttrsLocalizationTestData>()
            {
                new AttrsLocalizationTestData()
                {
                    Text = "啊实打实大苏打撒",
                    Name="aa",
                    Number =5000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
               new AttrsLocalizationTestData()
                {
                    Text = "啊实打实大苏打撒",
                    Name="啊实打实大苏打撒",
                    Number =6000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
               new AttrsLocalizationTestData()
                {
                    Text = "啊实打实速度大苏打撒",
                    Name="萨达萨达",
                    Number =6000,
                    Text2 = "突然他也让他人",
                    Text3 = "sadsad打发打发士大夫的"
                },
            });

### 导入 Demo

#### 导入特性（**ImporterAttribute**）：

- **HeaderRowIndex**：表头位置

#### 导入列头特性（**ImporterHeader**）：

+ **Name**：表头显示名称(不可为空)。
+ **Description**：表头添加注释。
+ **Author**：注释作者，默认值为“麦扣”。
+ **AutoTrim**：自动过滤空格，默认启用。
+ **FixAllSpace**：处理掉所有的空格，包括中间空格。默认false。
+ **ColumnIndex**：列索引，一般不建议设置。

#### 导入结果（ImportResult）：

+ **Data**：***IList\<T>***  导入的数据集合。
+ **RowErrors**：***IList<DataRowErrorInfo>*** 数据行错误。
+ **HasError**：***bool*** 是否存在导入错误。
+ **Exception**：异常信息
+ **TemplateErrors**：模板错误信息



---
#### Demo2-1 普通模板
##### 生成模板
![](./res/2-1.png "Demo2-1")

>
    public class ImportProductDto
    {
        /// <summary>
        /// 产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称")]
        public string Name { get; set; }
        /// <summary>
        /// 产品代码
        /// </summary>
        [ImporterHeader(Name = "产品代码")]
        public string Code { get; set; }
        /// <summary>
        /// 产品条码
        /// </summary>
        [ImporterHeader(Name = "产品条码")]
        public string BarCode { get; set; }
    }

##### 导入模板
![](./res/2-3.png "Demo2-3")
![](./res/2-4.png "Demo2-4")

---
#### Demo2-2 多数据类型
##### 生成模板
![](./res/2-2.png "Demo2-2")
>
    public class ImportProductDto
    {
        /// <summary>
        /// 产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称")]
        public string Name { get; set; }
        /// <summary>
        /// 产品代码
        /// </summary>
        [ImporterHeader(Name = "产品代码")]
        public string Code { get; set; }
        /// <summary>
        /// 产品条码
        /// </summary>
        [ImporterHeader(Name = "产品条码")]
        public string BarCode { get; set; }
        /// <summary>
        /// 客户Id
        /// </summary>
        [ImporterHeader(Name = "客户代码")]
        public long ClientId { get; set; }
        /// <summary>
        /// 产品型号
        /// </summary>
        [ImporterHeader(Name = "产品型号")]
        public string Model { get; set; }
        /// <summary>
        /// 申报价值
        /// </summary>
        [ImporterHeader(Name = "申报价值")]
        public double DeclareValue { get; set; }
        /// <summary>
        /// 货币单位
        /// </summary>
        [ImporterHeader(Name = "货币单位")]
        public string CurrencyUnit { get; set; }
        /// <summary>
        /// 品牌名称
        /// </summary>
        [ImporterHeader(Name = "品牌名称")]
        public string BrandName { get; set; }
        /// <summary>
        /// 尺寸
        /// </summary>
        [ImporterHeader(Name = "尺寸(长x宽x高)")]
        public string Size { get; set; }
        /// <summary>
        /// 重量
        /// </summary>
        [ImporterHeader(Name = "重量(KG)")]
        public double Weight { get; set; }

        /// <summary>
        /// 类型
        /// </summary>
        [ImporterHeader(Name = "类型")]
        public ImporterProductType Type { get; set; }
    
        /// <summary>
        /// 是否行
        /// </summary>
        [ImporterHeader(Name = "是否行")]
        public bool IsOk { get; set; }
    }

>
    public enum ImporterProductType
    {
        [Display(Name = "第一")]
        One,
        [Display(Name = "第二")]
        Two
    }
##### 导入模板
![](./res/2-5.png "Demo2-5")
![](./res/2-6.png "Demo2-6")

---
#### Demo2-3 数据验证
##### 生成模板
***必填项表头文本为红色***
![](./res/2-7.png "Demo2-7")

>
    public class ImportProductDto
    {
        /// <summary>
        /// 产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称",Description ="必填")]
        [Required(ErrorMessage = "产品名称是必填的")]
        public string Name { get; set; }
        /// <summary>
        /// 产品代码
        /// </summary>
        [ImporterHeader(Name = "产品代码", Description = "最大长度为8")]
        [MaxLength(8,ErrorMessage = "产品代码最大长度为8")]
        public string Code { get; set; }
        /// <summary>
        /// 产品条码
        /// </summary>
        [ImporterHeader(Name = "产品条码")]
        [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
        [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字")]
        public string BarCode { get; set; }
        /// <summary>
        /// 客户Id
        /// </summary>
        [ImporterHeader(Name = "客户代码")]
        public long ClientId { get; set; }
        /// <summary>
        /// 产品型号
        /// </summary>
        [ImporterHeader(Name = "产品型号")]
        public string Model { get; set; }
        /// <summary>
        /// 申报价值
        /// </summary>
        [ImporterHeader(Name = "申报价值")]
        public double DeclareValue { get; set; }
        /// <summary>
        /// 货币单位
        /// </summary>
        [ImporterHeader(Name = "货币单位")]
        public string CurrencyUnit { get; set; }
        /// <summary>
        /// 品牌名称
        /// </summary>
        [ImporterHeader(Name = "品牌名称")]
        public string BrandName { get; set; }
        /// <summary>
        /// 尺寸
        /// </summary>
        [ImporterHeader(Name = "尺寸(长x宽x高)")]
        public string Size { get; set; }
        /// <summary>
        /// 重量
        /// </summary>
        [ImporterHeader(Name = "重量(KG)")]
        public double Weight { get; set; }

        /// <summary>
        /// 类型
        /// </summary>
        [ImporterHeader(Name = "类型")]
        public ImporterProductType Type { get; set; }
    
        /// <summary>
        /// 是否行
        /// </summary>
        [ImporterHeader(Name = "是否行")]
        public bool IsOk { get; set; }
    }

>
    public enum ImporterProductType
    {
        [Display(Name = "第一")]
        One,
        [Display(Name = "第二")]
        Two
    }
##### 导入模板
![](./res/2-8.png "Demo2-8")
![](./res/2-9.png "Demo2-9")

#### Docker中使用

>
    # 安装libgdiplus库，用于Excel导出
    RUN apt-get update && apt-get install -y libgdiplus libc6-dev
    RUN ln -s /usr/lib/libgdiplus.so /usr/lib/gdiplus.dll

Dockerfile Demo 
>
    FROM microsoft/dotnet:2.2-aspnetcore-runtime AS base
    # 安装libgdiplus库，用于Excel导出
    RUN apt-get update && apt-get install -y libgdiplus libc6-dev
    RUN ln -s /usr/lib/libgdiplus.so /usr/lib/gdiplus.dll
    WORKDIR /app
    EXPOSE 80

    FROM microsoft/dotnet:2.2-sdk AS build
    WORKDIR /src
    COPY ["src/web/Admin.Host/Admin.Host.csproj", "src/web/Admin.Host/"]
    COPY ["src/web/Admin.Web.Core/Admin.Web.Core.csproj", "src/web/Admin.Web.Core/"]
    COPY ["src/application/Admin.Application/Admin.Application.csproj", "src/application/Admin.Application/"]
    COPY ["src/core/Magicodes.Admin.Core/Magicodes.Admin.Core.csproj", "src/core/Magicodes.Admin.Core/"]
    COPY ["src/data/Magicodes.Admin.EntityFrameworkCore/Magicodes.Admin.EntityFrameworkCore.csproj", "src/data/Magicodes.Admin.EntityFrameworkCore/"]
    COPY ["src/core/Magicodes.Admin.Core.Custom/Magicodes.Admin.Core.Custom.csproj", "src/core/Magicodes.Admin.Core.Custom/"]
    COPY ["src/application/Admin.Application.Custom/Admin.Application.Custom.csproj", "src/application/Admin.Application.Custom/"]
    RUN dotnet restore "src/web/Admin.Host/Admin.Host.csproj"
    COPY . .
    WORKDIR "/src/src/web/Admin.Host"
    RUN dotnet build "Admin.Host.csproj" -c Release -o /app
    
    FROM build AS publish
    RUN dotnet publish "Admin.Host.csproj" -c Release -o /app
    
    FROM base AS final
    WORKDIR /app
    COPY --from=publish /app .
    ENTRYPOINT ["dotnet", "Magicodes.Admin.Web.Host.dll"]
