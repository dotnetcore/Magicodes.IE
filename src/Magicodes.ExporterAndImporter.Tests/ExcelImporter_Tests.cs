using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models;
using Shouldly;
using System.IO;
using System.Threading.Tasks;
using Xunit;

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

        [Fact(DisplayName = "根据文件路径导入")]
        public async Task GenerateTemplate_Test()
        {
            //var filePath = Path.Combine(Directory.GetCurrentDirectory(), "testTemplate.xlsx");
            //if (File.Exists(filePath)) File.Delete(filePath);

            //var result = await Importer.GenerateTemplate<ImportProductDto>(filePath);
            //result.ShouldNotBeNull();
            //File.Exists(filePath).ShouldBeTrue();
            string filePath = @"D:\NewDeskTop\t2.xls";
            var result = Importer.Import<ProductDto>(filePath);
            result.ShouldNotBeNull();
        }

        [Fact(DisplayName = "根据文件流导入")]
        public async Task ImporterStream_Test()
        {
            string filePath = @"D:\NewDeskTop\t2.xlsx";
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var import = Importer.Import<ProductDto>(fs);
            import.ShouldNotBeNull();
        }
    }
}