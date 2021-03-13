// ======================================================================
// 
//           filename : ExportTestDataWithAttrs.cs
//           description :
// 
//           created by 雪雁 at  2019-11-05 20:02
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml.Table;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "测试忽略dto所有属性", TableStyle = TableStyles.Light10, AutoFitAllColumn = true)]
    public class ExportTestIgnoreAllColumns
    {
        [ExporterHeader(DisplayName = "忽略1", IsBold = true, IsIgnore = true)]
        public string IgnoreText1 { get; set; }

        [ExporterHeader(DisplayName = "忽略2", IsBold = true, IsIgnore = true)]
        public string IgnoreText2 { get; set; }

        [ExporterHeader(DisplayName = "忽略3", IsBold = true, IsIgnore = true)]
        public string IgnoreText3 { get; set; }
    }
}