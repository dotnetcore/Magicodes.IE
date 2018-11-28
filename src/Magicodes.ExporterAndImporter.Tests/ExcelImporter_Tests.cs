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

        [Fact(DisplayName = "µº»Î")]
        public async Task Importer_Test()
        {
            var import = await Importer.Import<ImportTestData>(
                @"D:\Coding\xin-lai.github\src\Magicodes.ExporterAndImporter.Tests\Models\Importer_test.xlsx");
            import.ShouldNotBeNull();
        }
    }
}
