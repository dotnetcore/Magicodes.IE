using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using OfficeOpenXml;
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

        [Fact(DisplayName = "DTO特性导出（测试格式化）")]
        public async Task AttrsExport_Test()
        {
            IExporter exporter = new ExcelExporter();

            var filePath = GetTestFilePath($"{nameof(AttrsExport_Test)}.csv");

            DeleteFile(filePath);

            var data = GenFu.GenFu.ListOf<ExportTestDataWithAttrs>(100);
            foreach (var item in data)
            {
                item.LongNo = 45875266524;
            }
            var result = await exporter.Export(filePath, data,EnumExportType.Csv);

            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                csv.Configuration.RegisterClassMap<AutoMap<ExportTestDataWithAttrs>>();
                var exportDatas = csv.GetRecords<ExportTestDataWithAttrs>().ToList();
                exportDatas.Count().ShouldBe(100);
                var exportData = exportDatas.FirstOrDefault();
                exportData.Time1.ToString().ShouldBeGreaterThanOrEqualTo(exportData.Time1.ToString("yyyy-MM-dd"));
                exportData.Time2.ToString().ShouldBeGreaterThanOrEqualTo(exportData.Time2?.ToString("yyyy-MM-dd HH:mm:ss"));
            }
        }

    }
}
