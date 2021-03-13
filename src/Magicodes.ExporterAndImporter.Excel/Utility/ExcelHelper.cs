// ======================================================================
// 
//           filename : ExcelHelper.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

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
        /// <summary>
        ///     创建Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="creator"></param>
        /// <returns></returns>
        public static ExportFileInfo CreateExcelPackage(string fileName, Action<ExcelPackage> creator)
        {
            var file = new ExportFileInfo(fileName,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            using (var excelPackage = new ExcelPackage())
            {
                creator(excelPackage);
                Save(excelPackage, file);
            }

            return file;
        }

        private static void Save(ExcelPackage excelPackage, ExportFileInfo file)
        {
            excelPackage.SaveAs(new FileInfo(file.FileName));
        }
    }
}