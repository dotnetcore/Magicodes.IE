using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class ExportTestDataWithColFontColor
    {
        [ExporterHeader(DisplayName = "姓名", IsBold = true, AutoCenterColumn = true)]
        public string Name { get; set; }

#if !NETCOREAPP2_1
        [ExporterHeader(DisplayName = "年龄", FontColor = System.Drawing.KnownColor.Red)]
#else
        [ExporterHeader(DisplayName = "年龄", FontColor = KnownColor.Red)]
#endif
        public int Age { get; set; }

    }
}