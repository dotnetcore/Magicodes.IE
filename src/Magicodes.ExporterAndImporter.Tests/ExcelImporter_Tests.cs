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
    public class ExcelImporter_Tests
    {
        public IImporter Importer = new ExcelImporter();

        [Fact(DisplayName = "导入")]
        public async Task Importer_Test()
        {
            var import = await Importer.Import<ImportProductDto>(
                @"G:\GitCodes\Magicodes.ExporterAndImporter\src\Magicodes.ExporterAndImporter.Tests\Models\testTemplate.xlsx");
            import.ShouldNotBeNull();
        }

        [Fact(DisplayName = "生成模板")]
        public async Task GenerateTemplate()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "testTemplate.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportProductDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }
    }
}
