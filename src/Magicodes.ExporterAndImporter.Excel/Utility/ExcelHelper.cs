using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    public static class ExcelHelper
    {
        /// <summary>
        ///     创建Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="creator"></param>
        /// <returns></returns>
        public static ExcelFileInfo CreateExcelPackage(string fileName, Action<ExcelPackage> creator)
        {
            var file = new ExcelFileInfo(fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            using (var excelPackage = new ExcelPackage())
            {
                creator(excelPackage);
                Save(excelPackage, file);
            }
            return file;
        }

        private static void Save(ExcelPackage excelPackage, ExcelFileInfo file)
        {
            excelPackage.SaveAs(new FileInfo(file.FileName));
        }
    }
}
