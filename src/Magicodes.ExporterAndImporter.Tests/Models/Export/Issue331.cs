using System;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "导出结果", TableStyle = OfficeOpenXml.Table.TableStyles.None)]
    public class Issue331
    {
        /// <summary>
        /// Time1
        /// </summary>
        [ExporterHeader(DisplayName = "Time1", Format = "yyyy-MM-dd HH:mm:ss")]
        public DateTime Time1 { get; set; }

        /// <summary>
        /// Time2
        /// </summary>
        [ExporterHeader(DisplayName = "Time2", Format = "yyyy-MM-dd")]
        public DateTime Time2 { get; set; }

        /// <summary>
        /// Time3
        /// </summary>
        [ExporterHeader(DisplayName = "Time3")]
        public DateTime Time3 { get; set; }

        /// <summary>
        /// Time4
        /// </summary>
        [ExporterHeader(DisplayName = "Time4")]
        public DateTime? Time4 { get; set; } 
        
        /// <summary>
        /// Time5
        /// </summary>
        [ExporterHeader(DisplayName = "Time5", Format = "HH:mm:ss")]
        public DateTime Time5 { get; set; } 
    }
}