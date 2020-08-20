using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class Issue140_IdCardExportDto
    {
        /// <summary>
        /// 身份证
        /// </summary>
        [ExporterHeader(DisplayName = "Text", Format = "@")]
        public string IdCard { get; set; }
    }
}
