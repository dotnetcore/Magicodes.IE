using Magicodes.ExporterAndImporter.Core;
using Magicodes.IE.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class ExportTestDataWithColFontColor
    {
        [ExporterHeader(DisplayName = "姓名", IsBold = true, AutoCenterColumn = true)]
        public string Name { get; set; }

        [ExporterHeader(fontColor: KnownColor.Red,  DisplayName = "年龄")] 
        public int Age { get; set; }
    }
}