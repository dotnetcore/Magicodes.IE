using System;
using System.Collections.Generic;
using System.Text;
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// 导入列头设置
    /// </summary>
    public class ImporterHeaderInfo
    {
        ///// <summary>
        ///// 列索引
        ///// </summary>
        //public int Index { get; set; }

        /// <summary>
        /// 列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 列属性
        /// </summary>
        public ImporterHeaderAttribute ExporterHeader { get; set; }
    }
}
