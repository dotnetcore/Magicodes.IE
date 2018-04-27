using Magicodes.ExporterAndImporter.Core;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// Excel导入
    /// </summary>
    public class ExcelImporter : IImporter
    {
        /// <summary>
        /// 导入为DataTable
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public Task<DataTable> Import(string filePath)
        {
            CheckImportFile(filePath);
            try
            {
                using (Stream stream = new FileStream(filePath, FileMode.Open))
                {
                    var dt = new DataTable();
                    using (var package = new ExcelPackage(stream))
                    {
                        //默认获取第一个
                        var sheet = package.Workbook.Worksheets[1];
                        //跳过第一行（列名）
                        var startRowIndx = sheet.Dimension.Start.Row + 1;
                        for (var rowIndex = startRowIndx; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                        {

                            var dr = dt.NewRow();
                            for (var colIndex = sheet.Dimension.Start.Column; colIndex <= sheet.Dimension.End.Column; colIndex++)
                            {
                                //第一次遍历时添加列
                                if (rowIndex == startRowIndx)
                                {
                                    dt.Columns.Add(colIndex.ToString(), Type.GetType("System.String"));
                                }

                                if (sheet.Cells[rowIndex, colIndex].Style.Numberformat.Format.IndexOf("yyyy") > -1
                                    && sheet.Cells[rowIndex, colIndex].Value != null)//处理日期时间格式
                                {

                                    dr[colIndex - 1] = sheet.Cells[rowIndex, colIndex].GetValue<DateTime>();
                                }
                                else
                                    dr[colIndex - 1] = (sheet.Cells[rowIndex, colIndex].Value ?? DBNull.Value);
                            }
                            dt.Rows.Add(dr);
                        }
                        return Task.FromResult(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ImportException("导入时出现未知错误!", ex);
            }

        }

        private static void CheckImportFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("文件路径不能为空!", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new ImportException("导入文件不存在!");
            }
        }

        public Task<IList<T>> Import<T>(string filePath) where T : class, new()
        {
            CheckImportFile(filePath);
            CreateExcelPackage(filePath, excelPackage =>
            {
                //导入定义
                var importer = GetImporterAttribute<T>();
                ExcelWorksheet worksheet = null;
                if (!string.IsNullOrWhiteSpace(importer.SheetName))
                {
                    worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(p => p.Name == importer.SheetName);
                    if (worksheet == null)
                    {
                        throw new ImportException("没有找到Sheet名称为 " + importer.SheetName + " 的Sheet!");
                    }
                }
                else
                    worksheet = excelPackage.Workbook.Worksheets[1];

                var propertyInfoList = new List<PropertyInfo>(typeof(T).GetProperties());
                //跳过第一行（列名）
                var startRowIndx = worksheet.Dimension.Start.Row + 1;
                for (var rowIndex = startRowIndx; rowIndex <= worksheet.Dimension.End.Row; rowIndex++)
                {
                    var dataItem = new T();
                    foreach (var propertyInfo in propertyInfoList)
                    {
                        var ColName = propertyInfo.Name;
                    }

                    for (var colIndex = worksheet.Dimension.Start.Column; colIndex <= worksheet.Dimension.End.Column; colIndex++)
                    {

                        if (worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format.IndexOf("yyyy") > -1
                            && worksheet.Cells[rowIndex, colIndex].Value != null)//处理日期时间格式
                        {

                            dr[colIndex - 1] = worksheet.Cells[rowIndex, colIndex].GetValue<DateTime>();
                        }
                        else
                            dr[colIndex - 1] = (worksheet.Cells[rowIndex, colIndex].Value ?? DBNull.Value);
                    }

                }

            });
        }

        /// <summary>
        ///     创建Excel
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="creator"></param>
        /// <returns></returns>
        protected void CreateExcelPackage(string fileName, Action<ExcelPackage> creator)
        {
            using (var excelPackage = new ExcelPackage())
            {
                creator(excelPackage);
            }
        }


        /// <summary>
        /// 获取导入全局定义
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private static ExcelImporterAttribute GetImporterAttribute<T>() where T : class
        {
            var exporterTableAttributes = (typeof(T).GetCustomAttributes(typeof(ExcelImporterAttribute), true) as ExcelImporterAttribute[]);
            if (exporterTableAttributes != null && exporterTableAttributes.Length > 0)
                return exporterTableAttributes[0];

            var exporterAttributes = (typeof(T).GetCustomAttributes(typeof(ImporterAttribute), true) as ImporterAttribute[]);

            if (exporterAttributes != null && exporterAttributes.Length > 0)
            {
                var export = exporterAttributes[0];
                return new ExcelImporterAttribute()
                {

                };
            }
            return null;
        }

    }
}
