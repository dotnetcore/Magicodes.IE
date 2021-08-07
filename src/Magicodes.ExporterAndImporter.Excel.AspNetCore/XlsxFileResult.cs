using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel.AspNetCore
{
    /// <summary>
    /// Excel文件ActionResult
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XlsxFileResult<T> : ActionResult where T : class, new()
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

            var response = context.HttpContext.Response;
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            context.HttpContext.Response.Headers.Add("Content-Disposition", new[] {
                "attachment; filename=" + (FileDownloadName?? Guid.NewGuid().ToString("N") + ".xlsx")
            });
            await fs.CopyToAsync(context.HttpContext.Response.Body);
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class XlsxFileResult : ActionResult
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileStream"></param>
        /// <param name="fileDownloadName"></param>
        public XlsxFileResult(Stream fileStream, string fileDownloadName = null)
        {
            FileStream = fileStream;
            FileDownloadName = fileDownloadName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="fileDownloadName"></param>

        public XlsxFileResult(byte[] bytes, string fileDownloadName = null)
        {
            FileStream = new MemoryStream(bytes);
            FileDownloadName = fileDownloadName;
        }


        public Stream FileStream { get; protected set; }
        public string FileDownloadName { get; protected set; }


        public async override Task ExecuteResultAsync(ActionContext context)
        {
            var response = context.HttpContext.Response;
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            context.HttpContext.Response.Headers.Add("Content-Disposition", new[] {
                "attachment; filename=" + (FileDownloadName?? Guid.NewGuid().ToString("N") + ".xlsx")
            });
            await FileStream.CopyToAsync(context.HttpContext.Response.Body);
        }
    }
}
