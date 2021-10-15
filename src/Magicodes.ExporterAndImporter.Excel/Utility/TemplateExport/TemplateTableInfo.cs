using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel.Utility.TemplateExport
{
    /// <summary>
    /// 模板表格信息
    /// </summary>
    public class TemplateTableInfo
    {
        /// <summary>
        /// 表格数据对象Key
        /// </summary>
        public string TableKey { get; set; }

        /// <summary>
        /// 原始开始行
        /// </summary>
        public int RawRowStart { get; set; }

        /// <summary>
        /// 新开始行
        /// </summary>
        public int NewRowStart { get; set; }

        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; set; }

        /// <summary>
        /// 写入器
        /// </summary>
        public IGrouping<string, IWriter> Writers { get; set; }
    }
}