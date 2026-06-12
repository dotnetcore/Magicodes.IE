using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Web;

namespace Magicodes.ExporterAndImporter.Excel.AspNetCore
{
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

            context.HttpContext.Response.Headers["Content-Disposition"] = 
                "attachment; filename=" + HttpUtility.UrlEncode(downloadFileName);
            await stream.CopyToAsync(context.HttpContext.Response.Body);
        }
    }
}
