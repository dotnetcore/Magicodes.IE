using Magicodes.ExporterAndImporter.Core;
using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Excel
{
    public class ExcelImporterAttribute : ImporterAttribute
    {
        public ExcelImporterAttribute()
        {
        }

        /// <summary>
        /// 指定Sheet名称(获取指定Sheet名称)
        /// 为空则自动获取第一个
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 是否标注错误（默认为true）
        /// </summary>
        public bool IsLabelingError { get; set; } = true;
    }
}
