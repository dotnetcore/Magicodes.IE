using Magicodes.ExporterAndImporter.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public  class ImportStudentAndPaymentLogDto
    {

        [ExcelImporter(SheetName = "1班导入数据")]
        public ImportStudentDto Class1Students { get; set; }

        [ExcelImporter(SheetName = "缴费数据")]
        public ImportPaymentLogDto Class2Students { get; set; }
    }
}
