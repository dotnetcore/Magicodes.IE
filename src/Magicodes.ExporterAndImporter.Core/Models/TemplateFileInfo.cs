// ======================================================================
// 
//           filename : TemplateFileInfo.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

namespace Magicodes.ExporterAndImporter.Core.Models
{
    public class TemplateFileInfo
    {
        public TemplateFileInfo()
        {
        }

        public TemplateFileInfo(string fileName, string fileType)
        {
            FileName = fileName;
            FileType = fileType;
        }

        public string FileName { get; set; }

        public string FileType { get; set; }
    }
}