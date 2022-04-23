using NPOI.XSSF.UserModel;
using System.IO;

namespace Magicodes.ExporterAndImporter.Excel.NPOI
{
    /// <summary>
    ///     扩展类
    /// </summary>
    public static class Extension
    {
        public static byte[] SaveToExcelWithXSSFWorkbook(this byte[] data)
        {
            //for excel compability
            using (var stream = new MemoryStream(data))
            {
                XSSFWorkbook wb = new XSSFWorkbook(stream);
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Write(ms);
                    return ms.ToArray();
                }
            }
        }
    }
}