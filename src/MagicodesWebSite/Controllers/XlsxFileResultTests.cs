using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Excel.AspNetCore;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MagicodesWebSite.Controllers
{
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
}
