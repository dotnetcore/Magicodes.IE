## Magicodes.IE.Excel.AspNetCore之快速导出Excel

## 前言

总是有很多朋友咨询Magicodes.IE如何基于ASP.NET Core导出Excel，出于从框架的体验和易用性的角度，Magicodes.IE决定对Excel的导出进行独立封装，以便于大家更易于使用，开箱即用。

***注意：Magicodes.IE是从框架的易用性和体验的角度对Excel导出进行了封装，但是希望大家先理解原理后再使用。***

### 1.安装包

```powershell
Install-Package Magicodes.IE.Excel.AspNetCore
```

### 2.引用命名空间

`using Magicodes.ExporterAndImporter.Excel.AspNetCore;`

## 3.直接使用XlsxFileResult

参考Demo如下所示：

```csharp
    [ApiController]
    [Route("api/[controller]")]
    public class XlsxFileResultTests : ControllerBase
    {
        /// <summary>
        /// 使用Byte数组导出Excel文件
        /// </summary>
        /// <returns></returns>
        [HttpGet("ByBytes")]
        public async Task<ActionResult> ByBytes()
        {
            //随机生成100条数据
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var exporter = new ExcelExporter();
            var bytes = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(list);
            //使用XlsxFileResult进行导出
            return new XlsxFileResult(bytes: bytes);
        }

        /// <summary>
        /// 使用流导出Excel文件
        /// </summary>
        /// <returns></returns>
        [HttpGet("ByStream")]
        public async Task<ActionResult> ByStream()
        {
            //随机生成100条数据
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var exporter = new ExcelExporter();
            var result = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(list);
            var fs = new MemoryStream(result);
            return new XlsxFileResult(stream: fs, fileDownloadName: "下载文件");
        }


        /// <summary>
        /// 使用泛型集合导出Excel文件
        /// </summary>
        /// <returns></returns>
        [HttpGet("ByList")]
        public async Task<ActionResult> ByList()
        {
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            return new XlsxFileResult<ExportTestDataWithAttrs>(data: list);
        }
    }
```

如上所示，引用 `Magicodes.IE.Excel.AspNetCore`之后，导出就会变得如此简单。值得注意的是：

1. 使用`XlsxFileResult`需引用包`Magicodes.IE.Excel.AspNetCore`
2. `XlsxFileResult`继承自`ActionResult`，目前支持**字节数组、流和泛型集合**为参数的Excel文件下载
3. 支持传递下载文件名，参数名`fileDownloadName`，如不传则自动生成唯一的文件名

### 核心实现

在`Magicodes.IE.Excel.AspNetCore`中，我们添加了自定义的`ActionResult`——`XlsxFileResult`，核心参考代码如下所示：

```csharp
    /// <summary>
    /// Excel文件ActionResult
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XlsxFileResult<T> : XlsxFileResultBase where T : class, new()
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="fileDownloadName"></param>
        public XlsxFileResult(ICollection<T> data, string fileDownloadName = null)
        {
            FileDownloadName = fileDownloadName;
            Data = data;
        }

        public string FileDownloadName { get; }
        public ICollection<T> Data { get; }

        public async override Task ExecuteResultAsync(ActionContext context)
        {
            var exporter = new ExcelExporter();
            var bytes = await exporter.ExportAsByteArray(Data);
            var fs = new MemoryStream(bytes);
            await DownloadExcelFileAsync(context, fs, FileDownloadName);
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class XlsxFileResult : XlsxFileResultBase
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="fileDownloadName"></param>
        public XlsxFileResult(Stream stream, string fileDownloadName = null)
        {
            Stream = stream;
            FileDownloadName = fileDownloadName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="fileDownloadName"></param>

        public XlsxFileResult(byte[] bytes, string fileDownloadName = null)
        {
            Stream = new MemoryStream(bytes);
            FileDownloadName = fileDownloadName;
        }


        public Stream Stream { get; protected set; }
        public string FileDownloadName { get; protected set; }


        public async override Task ExecuteResultAsync(ActionContext context)
        {
            await DownloadExcelFileAsync(context, Stream, FileDownloadName);
        }
    }

    /// <summary>
    /// 基类
    /// </summary>
    public class XlsxFileResultBase : ActionResult
    {
        /// <summary>
        /// 下载Excel文件
        /// </summary>
        /// <param name="context"></param>
        /// <param name="stream"></param>
        /// <param name="downloadFileName"></param>
        /// <returns></returns>
        protected virtual async Task DownloadExcelFileAsync(ActionContext context,
                                                            Stream stream,
                                                            string downloadFileName)
        {
            var response = context.HttpContext.Response;
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            if (downloadFileName == null)
            {
                downloadFileName = Guid.NewGuid().ToString("N") + ".xlsx";
            }

            if (string.IsNullOrEmpty(Path.GetExtension(downloadFileName)))
            {
                downloadFileName += ".xlsx";
            }

            context.HttpContext.Response.Headers.Add("Content-Disposition", new[] {
                "attachment; filename=" +HttpUtility.UrlEncode(downloadFileName)
            });
            await stream.CopyToAsync(context.HttpContext.Response.Body);
        }
    }
```

欢迎大家多多PR并且前来解锁更多玩法。

## 知识点总结

敲黑板，麻烦有空帮点点，为了德玛西亚：

[ASP.NET Core Web API 中控制器操作的返回类型 | Microsoft Docs](https://docs.microsoft.com/zh-cn/aspnet/core/web-api/action-return-types?view=aspnetcore-5.0&WT.mc_id=DT-MVP-5004079)

## Reference

https://github.com/dotnetcore/Magicodes.IE

## 最后

有兴趣有精力的朋友可以帮忙PR一下单元测试，由于精力有限，先手测了，参考：

[ASP.NET Core 中的测试控制器逻辑 | Microsoft Docs](https://docs.microsoft.com/zh-cn/aspnet/core/mvc/controllers/testing?view=aspnetcore-5.0&WT.mc_id=DT-MVP-5004079)

写个功能几分钟到十几分钟，码个文档要半天，结束。









