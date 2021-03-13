// ======================================================================
// 
//           filename : Writer.cs
//           description :
// 
//           created by 雪雁 at  -- 
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

namespace Magicodes.ExporterAndImporter.Excel.Utility.TemplateExport
{
    /// <summary>
    /// 写入器
    /// </summary>
    public class Writer : IWriter
    {
        /// <summary>
        /// 地址
        /// </summary>
        public string TplAddress { get; set; }

        /// <summary>
        /// 单元格原始字符串
        /// </summary>
        public string CellString { get; set; }

        /// <summary>
        /// 写入的字符串
        /// </summary>
        public string WriteString { get; set; }

        /// <summary>
        /// 写入器类型
        /// </summary>
        public WriterTypes WriterType { get; set; }

        /// <summary>
        /// 表格数据对象Key
        /// </summary>
        public string TableKey { get; set; }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 列号
        /// </summary>
        public int ColIndex { get; set; }
    }
}