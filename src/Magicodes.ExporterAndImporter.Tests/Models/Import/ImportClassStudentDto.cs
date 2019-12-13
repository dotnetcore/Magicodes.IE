using Magicodes.ExporterAndImporter.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public class ImportClassStudentDto
    {

        [ExcelImporter(SheetName = "1班导入数据")]
        public ImportStudentDto Class1Students { get; set; }

        [ExcelImporter(SheetName = "2班导入数据")]
        public ImportStudentDto Class2Students { get; set; }

    }
}
