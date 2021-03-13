using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Html;
using Magicodes.ExporterAndImporter.Pdf;
using Magicodes.ExporterAndImporter.Word;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Extensions
{
    public class MagicodesBase
    {
        public async Task<string> ReadResponseBodyStreamAsync(Stream bodyStream)
        {
            bodyStream.Seek(0, SeekOrigin.Begin);
            var responseBody = await new StreamReader(bodyStream).ReadToEndAsync();
            bodyStream.Seek(0, SeekOrigin.Begin);
            return responseBody;
        }
        public static DataTable ToDataTable(string json)
        {
            return JsonConvert.DeserializeObject<DataTable>(json);
        }
        public async Task<bool> HandleSuccessfulReqeustAsync(HttpContext context, object body, Type type, string tplPath)
        {
            var contentType = "";
            string filename = DateTime.Now.ToString("yyyyMMddHHmmss");
            byte[] result = null;
            switch (context.Request.Headers["Magicodes-Type"])
            {
                case HttpContentMediaType.XLSXHttpContentMediaType:
                    filename += ".xlsx";
                    var dt = ToDataTable(body?.ToString());
                    contentType = HttpContentMediaType.XLSXHttpContentMediaType;
                    var exporter = new ExcelExporter();
                    result = await exporter.ExportAsByteArray(dt, type);
                    break;
                case HttpContentMediaType.PDFHttpContentMediaType:
                    filename += ".pdf";
                    contentType = HttpContentMediaType.PDFHttpContentMediaType;
                    IExportFileByTemplate pdfexporter = new PdfExporter();
                    var tpl = await File.ReadAllTextAsync(tplPath);
                    var obj = JsonConvert.DeserializeObject(body.ToString(), type);
                    result = await pdfexporter.ExportBytesByTemplate(obj, tpl, type);
                    break;
                case HttpContentMediaType.HTMLHttpContentMediaType:
                    filename += ".html";
                    contentType = HttpContentMediaType.HTMLHttpContentMediaType;
                    IExportFileByTemplate htmlexporter = new HtmlExporter();
                    result = await htmlexporter.ExportBytesByTemplate(JsonConvert.DeserializeObject(body.ToString(), type), await File.ReadAllTextAsync(tplPath), type);
                    break;
                case HttpContentMediaType.DOCXHttpContentMediaType:
                    filename += ".docx";
                    contentType = HttpContentMediaType.DOCXHttpContentMediaType;
                    IExportFileByTemplate docxexporter = new WordExporter();
                    result = await docxexporter.ExportBytesByTemplate(JsonConvert.DeserializeObject(body.ToString(), type), await File.ReadAllTextAsync(tplPath), type);
                    break;
            }
            if (contentType != "")
            {
                context.Response.Headers.Add("Content-Disposition", $"attachment;filename={filename}");
                context.Response.ContentType = contentType;
                if (result != null) await context.Response.Body.WriteAsync(result, 0, result.Length);
            }
            else
            {
                return false;
            }
            return true;

        }
    }

}
