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
    /// <summary>
    /// 导出文件信息
    /// </summary>
    public class ExportFileInfo
    {
        /// <summary>
        ///
        /// </summary>
        public ExportFileInfo()
        {
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileType"></param>
        public ExportFileInfo(string fileName, string fileType)
        {
            FileName = fileName;
            FileType = fileType;
        }

        /// <summary>
        /// 文件名（路径）
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件Mine类型
        /// </summary>
        public string FileType { get; set; }
    }
}