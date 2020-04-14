// ======================================================================
// 
//           filename : ExportTestDataWithAttrs.cs
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
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System;
using CsvHelper.Configuration.Attributes;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{

    [ExcelExporter(Name = "测试", TableStyle = "Light10", AutoFitAllColumn = true)]
    public class ExportByHeaders_TestDto
    {
        public string Text { get; set; }
        public string Text2 { get; set; }
        
        public string Text3 { get; set; }
        
        public decimal Number { get; set; }
        
        public string Name { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        public DateTime Time1 { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        public DateTime? Time2 { get; set; }

        public DateTime Time3 { get; set; }

        public DateTime Time4 { get; set; }

        /// <summary>
        /// 长数值测试
        /// </summary>
        public long LongNo { get; set; }
    }
}