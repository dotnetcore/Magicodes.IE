using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;
using System;
using System.IO;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    /// <summary>
    /// Excel辅助类
    /// </summary>
    public static class ExcelHelper
    {
        private const string _filetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <summary>
        ///     创建Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="creator"></param>
        /// <returns></returns>
        public static ExcelFileInfo CreateExcelPackage(string fileName, Action<ExcelPackage> creator)
        {
            var file = new ExcelFileInfo(fileName, _filetype);

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