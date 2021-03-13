// ======================================================================
//
//           filename : ImportStudentDto.cs
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

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public class MergeRowsImportDto
    {
        [ImporterHeader(Name = "学号")]
        public long No { get; set; }

        [ImporterHeader(Name = "姓名")]
        public string Name { get; set; }

        [ImporterHeader(Name = "性别")]
        public string Sex { get; set; }
    }
}