using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{

    [ExcelExporter(AutoCenter = true)]
    public class ExportTestDataWithAutoCenter
    {
        [ExporterHeader(DisplayName = "姓名", IsBold = true)]
        public string Name { get; set; }

        [ExporterHeader(DisplayName = "年龄")]
        public int Age { get; set; }

    }

    public class ExportTestDataWithColAutoCenter
    {
        [ExporterHeader(DisplayName = "姓名", IsBold = true, AutoCenterColumn = true)]
        public string Name { get; set; }

        [ExporterHeader(DisplayName = "年龄")]
        public int Age { get; set; }

    }
}
