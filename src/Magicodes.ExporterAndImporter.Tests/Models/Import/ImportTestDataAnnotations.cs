using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    [ExcelImporter(SheetName = "SheetName")]
    public class ImportTestDataAnnotations
    {
        /// <summary>
        ///     测试优先级
        /// </summary>
        [Display(Name = "Custom列")]
        [ImporterHeader(Name = "Custom列1")]
        public string Name { get; set; }
        [Display(Name = "列2", Order = 1)]
        public string Name1 { get; set; }
        [ImporterHeader(Name = "Time1")]
        public DateTime Time { get; set; }

        [ImporterHeader(Name = "Time2", IsIgnore = true)]
        [IEIgnoreAttribute]
        public DateTime Time2 { get; set; }


    }
}