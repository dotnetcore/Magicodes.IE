using Magicodes.ExporterAndImporter.Core;
using SixLabors.ImageSharp;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class ExportTestDataWithColFontColor
    {
        [ExporterHeader(DisplayName = "姓名", IsBold = true, AutoCenterColumn = true)]
        public string Name { get; set; }

        [ExporterHeader(DisplayName = "年龄")] 
        public int Age { get; set; }

    }
}