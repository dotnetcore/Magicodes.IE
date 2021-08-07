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
}
