using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    /// #597: CSV 长数字导出测试 DTO。
    /// </summary>
    public class CsvLongNumberExportDto
    {
        [ExporterHeader(DisplayName = "序号")]
        public int Id { get; set; }

        [ExporterHeader(DisplayName = "卡号", Format = "@")]
        public string CardNumber { get; set; }

        [ExporterHeader(DisplayName = "手机号", Format = "0")]
        public string PhoneNumber { get; set; }
    }
}