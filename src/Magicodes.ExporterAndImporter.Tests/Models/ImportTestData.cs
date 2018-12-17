using System;
using System.Collections.Generic;
using System.Text;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.ExporterAndImporter.Tests.Models
{
    //[ExcelImporter(SheetName = "Sheet1")]
    public class ImportTestData
    {
        [ImporterHeaderAttribute(Name ="1")]
        public string Name1 { get; set; }
        [ImporterHeaderAttribute(Name ="2")]
        public string Name2 { get; set; }
        [ImporterHeaderAttribute(Name ="3")]
        public string Name3 { get; set; }
        [ImporterHeaderAttribute(Name ="4")]
        public string Name4 { get; set; }
    }
}
