# Magicodes.ExporterAndImporter

导入导出通用库

### 特点

- 封装导入导出业务,目前仅支持 Excel,后续将支持 CSV 以及其他业务
- 配置特性即可控制相关逻辑和显示结果，无需修改逻辑代码
- 推荐配合 DTO 使用
- 导出支持列头自定义处理以便支持多语言等场景
- 导出支持文本自定义过滤或处理
- 导入支持自动根据 DTO 生成导入模板及模板验证
- 导入支持数据验证逻辑
- 导入支持数据下拉选择
- 导入支持注释添加

### 相关官方Nuget包

| 名称     |      Nuget      |
|----------|:-------------:|
| Magicodes.IE.Core  |  [![NuGet](https://buildstats.info/nuget/Magicodes.IE.Core)](https://www.nuget.org/packages/Magicodes.IE.Core) |
| Magicodes.IE.Excel |    [![NuGet](https://buildstats.info/nuget/Magicodes.IE.Excel)](https://www.nuget.org/packages/Magicodes.IE.Excel)   |


### 更新历史

#### 2019.9.11

- 导入支持自动去除前后空格，默认启用，可以针对列进行关闭，具体见AutoTrim设置
- 导入Dto的字段允许不设置ImporterHeader，支持通过DisplayAttribute特性获取列名
- 导入的Excel移除对Sheet名称的约束，默认获取第一个Sheet
- 完善导入的单元测试

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
>导入特性（**ImporterHeader**）：

+ **Name**：表头显示名称(不可为空)。

+ **Description**：表头添加注释。

+ **Author**：注释作者，默认值为X.M。

>导入结果（**ImportModel\<T>**）：

+ **Data**：***IList\<T>***  导入的数据集合。

+ **ValidationResults**：***IList\<ValidationResultModel>*** 数据验证结果。

+ **HasValidTemplate**：***bool*** 模板验证是否通过。

>数据验证结果（**ValidationResultModel**）：

+ **Index**：***int***  错误数据所在行。

+ **Errors**：***IDictionary<string, string>*** 整个Excel错误集合。目前仅支持数据验证错误。

+ **FieldErrors**：***IDictionary<string, string>*** 数据验证错误。

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
