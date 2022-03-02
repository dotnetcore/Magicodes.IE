// ======================================================================
//
//           filename : ExcelExporter_Tests.cs
//           description :
//
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
//
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using System.ComponentModel;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class DisplayName_GET_TestDto
    {
        [ExporterHeader(IsAutoFit = false)]
        [DisplayName("备注")]
        public string Remark { get; set; }

    }
}