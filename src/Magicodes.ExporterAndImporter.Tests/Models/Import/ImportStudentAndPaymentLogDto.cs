using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public class ImportStudentAndPaymentLogDto
    {

        [ExcelImporter(SheetName = "1班导入数据")]
        public ImportStudentDto Class1Students { get; set; }

#if NET461
        [ExcelImporter(SheetIndex = 2)]
#else
        [ExcelImporter(SheetIndex = 1)]
#endif
        public ImportPaymentLogDto Class2Students { get; set; }
    }
}
