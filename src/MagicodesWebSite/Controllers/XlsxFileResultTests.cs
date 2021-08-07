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
        [HttpGet("ByBytes")]
        public async Task<ActionResult> ByBytes()
        {
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var exporter = new ExcelExporter();
            var result = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(list);
            return new XlsxFileResult(result);
        }

        [HttpGet("ByStream")]
        public async Task<ActionResult> ByStream()
        {
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            var exporter = new ExcelExporter();
            var result = await exporter.ExportAsByteArray<ExportTestDataWithAttrs>(list);
            var fs = new MemoryStream(result);
            return new XlsxFileResult(fs);
        }

        [HttpGet("ByList")]
        public async Task<ActionResult> ByList()
        {
            var list = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            return new XlsxFileResult<ExportTestDataWithAttrs>(list);
        }
    }
}
