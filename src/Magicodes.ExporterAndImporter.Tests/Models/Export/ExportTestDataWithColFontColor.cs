using Magicodes.ExporterAndImporter.Core;
using Magicodes.IE.Core;
using SixLabors.ImageSharp;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class ExportTestDataWithColFontColor
    {
        [ExporterHeader(DisplayName = "姓名", IsBold = true, AutoCenterColumn = true)]
        public string Name { get; set; }

        [ExporterHeader(DisplayName = "年龄", FontColor = KnownColor.Red)] 
        public int Age { get; set; }

    }
}