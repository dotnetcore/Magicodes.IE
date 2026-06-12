# Magicodes.IE.Excel.AspNetCore Quick Export Excel

## Preface

Many friends often ask how to export Excel based on ASP.NET Core using Magicodes.IE. From the perspective of framework experience and ease of use, Magicodes.IE decided to independently package Excel export to make it easier to use and ready to use out of the box.

***Note: Magicodes.IE has packaged Excel export from the perspective of framework ease of use and experience, but we hope everyone understands the principles before using it.***

### 1. Install Package

```powershell
Install-Package Magicodes.IE.Excel.AspNetCore
```

### 2. Reference Namespace

`using Magicodes.ExporterAndImporter.Excel.AspNetCore;`

## 3. Direct Use of XlsxFileResult

Reference Demo is shown below:

```csharp
    [ApiController]
    [Route("api/[controller]")]
    public class XlsxFileResultTests : ControllerBase
    {
        /// <summary>
        /// Export Excel file using Byte array
        /// </summary>
        /// <returns></returns>
        [HttpGet("ByBytes")]
        public async Task<ActionResult> ByBytes()
        {
            //Randomly generate 100 pieces of data
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var exporter = new ExcelExporter();
            var bytes = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(list);
            //Use XlsxFileResult to export
            return new XlsxFileResult(bytes: bytes);
        }

        /// <summary>
        /// Export Excel file using stream
        /// </summary>
        /// <returns></returns>
        [HttpGet("ByStream")]
        public async Task<ActionResult> ByStream()
        {
            //Randomly generate 100 pieces of data
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var exporter = new ExcelExporter();
            var result = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(list);
            var fs = new MemoryStream(result);
            return new XlsxFileResult(stream: fs, fileDownloadName: "Download File");
        }


        /// <summary>
        /// Export Excel file using generic collection
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

As shown above, after referencing `Magicodes.IE.Excel.AspNetCore`, exporting becomes so simple. It is worth noting:

1. Using `XlsxFileResult` requires referencing the package `Magicodes.IE.Excel.AspNetCore`
2. `XlsxFileResult` inherits from `ActionResult` and currently supports Excel file download with **byte array, stream and generic collection** as parameters
3. Supports passing download file name, parameter name `fileDownloadName`, if not passed, a unique file name will be automatically generated

### Core Implementation

In `Magicodes.IE.Excel.AspNetCore`, we added a custom `ActionResult` - `XlsxFileResult`, the core reference code is shown below:

```csharp
    /// <summary>
    /// Excel file ActionResult
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
    /// Base class
    /// </summary>
    public class XlsxFileResultBase : ActionResult
    {
        /// <summary>
        /// Download Excel file
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

Welcome everyone to contribute PRs and unlock more features.

## Knowledge Summary

Key points, please help click when you have time, for Demacia:

[ASP.NET Core Web API controller action return types | Microsoft Docs](https://docs.microsoft.com/en-us/aspnet/core/web-api/action-return-types?view=aspnetcore-5.0&WT.mc_id=DT-MVP-5004079)

## Reference

https://github.com/dotnetcore/Magicodes.IE

## Finally

Friends who are interested and have energy can help PR unit tests. Due to limited energy, manual testing was done first. Reference:

[Test controller logic in ASP.NET Core | Microsoft Docs](https://docs.microsoft.com/en-us/aspnet/core/mvc/controllers/testing?view=aspnetcore-5.0&WT.mc_id=DT-MVP-5004079)

Writing a feature takes a few minutes to more than ten minutes, but writing documentation takes half a day. That's it.

**Magicodes.IE: Import and export general library, support Dto import and export, template export, fancy export and dynamic export, support Excel, Csv, Word, Pdf and Html.**

- Github：<https://github.com/dotnetcore/Magicodes.IE>
- Gitee (manually synced, not maintained)：<https://gitee.com/magicodes/Magicodes.IE>

**Related libraries will continue to be updated, and there may be slight differences from this tutorial in terms of functional experience. Please refer to the specific code, version logs, and unit test examples.**

