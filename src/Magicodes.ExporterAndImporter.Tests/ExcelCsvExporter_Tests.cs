using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Shouldly;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelCsvExporter_Tests: TestBase
    {

        [Fact(DisplayName = "大量数据导出Excel")]
        public async Task Export_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(Export_Test) + ".csv");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ExportTestData>(100000),EnumExportType.Csv);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }


        [Fact(DisplayName = "Dto导出")]
        public async Task ExportAsByteArray_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath= GetTestFilePath($"{nameof(ExportAsByteArray_Test)}.csv");
            DeleteFile(filePath);
            var result = await exporter.ExportAsByteArray(GenFu.GenFu.ListOf<ExportTestData>(), EnumExportType.Csv);
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            File.WriteAllBytes(filePath, result);
            File.Exists(filePath).ShouldBeTrue();
        }

    }
}
