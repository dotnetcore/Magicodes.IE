using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.Benchmarks.Models
{
    [ExcelExporter(Name = "测试", TableStyle = "Light10", AutoFitAllColumn = true,MaxRowNumberOnASheet = 50000)]
    public class ExportTestDataWithAttrs
    {
        [ExporterHeader(DisplayName = "数值", IsBold = true)]
        public int Age { get; set; }
        [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
        public string Name { get; set; }
        [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
        public string Text3 { get; set; }
    }
}
