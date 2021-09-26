using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "导出结果", TableStyle = OfficeOpenXml.Table.TableStyles.None)]
    public class Issue337
    {
        /// <summary>
        /// 名称
        /// </summary>
        [ExporterHeader(DisplayName = "姓名")]
        public string Name { get; set; }

        /// <summary>
        /// 性别
        /// </summary>
        [ExporterHeader(DisplayName = "性别")]
        public string Gender { get; set; }

        /// <summary>
        /// 是否校友
        /// </summary>
        [ExporterHeader(DisplayName = "是否校友")]
        [ValueMapping("是", true)]
        [ValueMapping("否", false)]
        public bool? IsAlumni { get; set; }

    }
}
