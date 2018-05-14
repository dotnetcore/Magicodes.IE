using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models;
using Xunit;
using System.IO;
using Shouldly;
using Magicodes.ExporterAndImporter.Excel.Builder;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelImporter_Tests
    {
        public IImporter Importer = new ExcelImporter();

        [Fact(DisplayName = "µ¼ÈëÎªDataTable")]
        public async Task Export_Test()
        {
            //var filePath = Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
            //if (File.Exists(filePath)) File.Delete(filePath);

            //var result = await Exporter.Export(filePath, new List<ExportTestData>()
            //{
            //    new ExportTestData()
            //    {
            //        Name1 = "1",
            //        Name2 = "test",
            //        Name3 = "12",
            //        Name4 = "11",
            //    },
            //    new ExportTestData()
            //    {
            //        Name1 = "1",
            //        Name2 = "test",
            //        Name3 = "12",
            //        Name4 = "11",
            //    }
            //});
            //result.ShouldNotBeNull();
            //File.Exists(filePath).ShouldBeTrue();
        }

        
    }
}
