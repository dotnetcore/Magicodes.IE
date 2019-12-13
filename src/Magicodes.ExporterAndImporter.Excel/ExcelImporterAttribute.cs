// ======================================================================
// 
//           filename : ExcelImporterAttribute.cs
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
using System;

namespace Magicodes.ExporterAndImporter.Excel
{

    [AttributeUsage(AttributeTargets.Class|AttributeTargets.Property)]
    public class ExcelImporterAttribute : ImporterAttribute
    {
        /// <summary>
        ///     指定Sheet名称(获取指定Sheet名称)
        ///     为空则自动获取第一个
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        ///     截止读取的列数（从1开始，如果已设置，则将支持空行以及特殊列）
        /// </summary>
        public int? EndColumnCount { get; set; }

        /// <summary>
        ///     是否标注错误（默认为true）
        /// </summary>
        public bool IsLabelingError { get; set; } = true;
    }
}