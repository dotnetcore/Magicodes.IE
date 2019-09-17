namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// Excel文件信息
    /// </summary>
    public class ExcelFileInfo
    {
        /// <summary>
        /// 文件名称
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件类型
        /// </summary>
        public string FileType { get; set; }

        /// <summary>
        /// 有参构造函数
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileType"></param>
        public ExcelFileInfo(string fileName, string fileType)
        {
            FileName = fileName;
            FileType = fileType;
        }
    }
}