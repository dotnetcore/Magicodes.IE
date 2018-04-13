using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models;
using Xunit;
using System.IO;
using Shouldly;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelExporter_Tests
    {
        public IExporter Exporter = new ExcelExporter();

        [Fact]
        public async Task Export_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
            var result = await Exporter.Export(filePath, new List<ExportTestData>()
            {
                new ExportTestData()
                {
                    Name1 = "1",
                    Name2 = "test",
                    Name3 = "12",
                    Name4 = "11",
                }
            });
            result.ShouldNotBeNull();

        }
    }
}
