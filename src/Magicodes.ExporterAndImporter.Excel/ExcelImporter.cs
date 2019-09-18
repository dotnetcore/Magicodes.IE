using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Magicodes.ExporterAndImporter.Core.Extension;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     Excel导入
    /// </summary>
    public class ExcelImporter : IImporter
    {
        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">文件名必须填写! - fileName</exception>
        public Task<TemplateFileInfo> GenerateTemplate<T>(string fileName) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("文件名必须填写!", nameof(fileName));

            var fileInfo =
                ExcelHelper.CreateExcelPackage(fileName, excelPackage => { StructureExcel<T>(excelPackage); });
            return Task.FromResult(fileInfo);
        }

        /// <summary>
        ///     生成Excel导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>二进制字节</returns>
        public Task<byte[]> GenerateTemplateByte<T>() where T : class
        {
            using (var excelPackage = new ExcelPackage())
            {
                StructureExcel<T>(excelPackage);
                return Task.FromResult(excelPackage.GetAsByteArray());
            }
        }

        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(string filePath) where T : class, new()
        {
            var result = new ImportResult<T>();
            try
            {
                var importerAttribute = typeof(T).GetAttribute<ExcelImporterAttribute>(true) ?? new ExcelImporterAttribute();

                CheckImportFile(filePath);
                result.RowErrors = new List<DataRowErrorInfo>();
                using (Stream stream = new FileStream(filePath, FileMode.Open))
                {
                    using (var excelPackage = new ExcelPackage(stream))
                    {
                        #region 检查模板
                        var tplErrors = ParseTemplate<T>(excelPackage, importerAttribute, out var columnHeaders);
                        result.TemplateErrors = tplErrors;
                        if (result.HasError)
                        {
                            return Task.FromResult(result);
                        }
                        #endregion

                        result.Data = ParseData<T>(importerAttribute, excelPackage, columnHeaders, result.RowErrors);
                        for (var i = 0; i < result.Data.Count; i++)
                        {
                            var isValid = ValidatorHelper.TryValidate(result.Data[i], out var validationResults);
                            if (!isValid)
                            {
                                var dataRowError = new DataRowErrorInfo
                                {
                                    RowIndex = i + 1
                                };
                                foreach (var validationResult in validationResults)
                                {
                                    var key = validationResult.MemberNames.First();
                                    var column = columnHeaders.FirstOrDefault(a => a.PropertyName == key);
                                    if (column != null)
                                    {
                                        key = column.ExporterHeader.Name;
                                    }
                                    var value = validationResult.ErrorMessage;
                                    if (dataRowError.FieldErrors.ContainsKey(key))
                                        dataRowError.FieldErrors[key] += ("," + value);
                                    else
                                        dataRowError.FieldErrors.Add(key, value);
                                }
                                result.RowErrors.Add(dataRowError);
                            }
                        }
                        return Task.FromResult(result);
                    }
                }
            }
            catch (Exception ex)
            {
                result.Exception = ex;
            }
            return Task.FromResult(result);
        }

        /// <summary>
        /// 检查导入文件路劲
        /// </summary>
        /// <exception cref="ArgumentException">文件路径不能为空! - filePath</exception>
        private static void CheckImportFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("文件路径不能为空!", nameof(filePath));

            //TODO:在Docker容器中存在文件路径找不到问题，暂时先注释掉
            //if (!File.Exists(filePath))
            //{
            //    throw new ImportException("导入文件不存在!");
            //}
        }

        /// <summary>
        /// 解析头部
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <exception cref="ArgumentException">导入实体没有定义ImporterHeader属性</exception>
        private static bool ParseImporterHeader<T>(out List<ImporterHeaderInfo> importerHeaderList,
            out Dictionary<int, IDictionary<string, int>> enumColumns, out List<int> boolColumns)
        {
            importerHeaderList = new List<ImporterHeaderInfo>();
            enumColumns = new Dictionary<int, IDictionary<string, int>>();
            boolColumns = new List<int>();
            var objProperties = typeof(T).GetProperties();
            if (objProperties == null || objProperties.Length == 0)
                return false;
            for (var i = 0; i < objProperties.Length; i++)
            {
                //TODO:简化并重构
                //如果不设置，则自动使用默认定义
                var importerHeaderAttribute =
                    (objProperties[i].GetCustomAttributes(typeof(ImporterHeaderAttribute), true) as
                        ImporterHeaderAttribute[])?.FirstOrDefault() ?? new ImporterHeaderAttribute()
                        {
                            Name = objProperties[i].GetDisplayName() ?? objProperties[i].Name
                        };

                if (string.IsNullOrWhiteSpace(importerHeaderAttribute.Name))
                {
                    importerHeaderAttribute.Name = objProperties[i].GetDisplayName() ?? objProperties[i].Name;
                }

                importerHeaderList.Add(new ImporterHeaderInfo
                {
                    IsRequired = objProperties[i].IsRequired(),
                    PropertyName = objProperties[i].Name,
                    ExporterHeader = importerHeaderAttribute
                });
                if (objProperties[i].PropertyType.BaseType?.Name.ToLower() == "enum")
                    enumColumns.Add(i + 1, objProperties[i].PropertyType.GetEnumDisplayNames());
                if (objProperties[i].PropertyType == typeof(bool)) boolColumns.Add(i + 1);
            }

            return true;
        }

        /// <summary>
        ///     构建Excel模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        private static void StructureExcel<T>(ExcelPackage excelPackage) where T : class
        {
            var worksheet = excelPackage.Workbook.Worksheets.Add(typeof(T).GetDisplayName() ?? "导入数据");
            if (!ParseImporterHeader<T>(out var columnHeaders, out var enumColumns, out var boolColumns)) return;

            //设置列头
            for (var i = 0; i < columnHeaders.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = columnHeaders[i].ExporterHeader.Name;
                if (!string.IsNullOrWhiteSpace(columnHeaders[i].ExporterHeader.Description))
                {
                    worksheet.Cells[1, i + 1].AddComment(columnHeaders[i].ExporterHeader.Description, columnHeaders[i].ExporterHeader.Author);
                }
                //如果必填，则列头标红
                if (columnHeaders[i].IsRequired)
                {
                    worksheet.Cells[1, i + 1].Style.Font.Color.SetColor(Color.Red);
                }
            }
            worksheet.Cells.AutoFitColumns();
            worksheet.Cells.Style.WrapText = true;
            worksheet.Cells[worksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[worksheet.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[worksheet.Dimension.Address].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[worksheet.Dimension.Address].Style.Fill.BackgroundColor.SetColor(Color.DarkSeaGreen);

            //枚举处理
            foreach (var enumColumn in enumColumns)
            {
                var range = ExcelCellBase.GetAddress(1, enumColumn.Key, ExcelPackage.MaxRows, enumColumn.Key);
                var dataValidations = worksheet.DataValidations.AddListValidation(range);
                foreach (var displayName in enumColumn.Value) dataValidations.Formula.Values.Add(displayName.Key);
            }

            //Bool类型处理
            foreach (var boolColumn in boolColumns)
            {
                var range = ExcelCellBase.GetAddress(1, boolColumn, ExcelPackage.MaxRows, boolColumn);
                var dataValidations = worksheet.DataValidations.AddListValidation(range);
                dataValidations.Formula.Values.Add("是");
                dataValidations.Formula.Values.Add("否");
            }
        }

        /// <summary>
        ///     解析模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private IList<TemplateErrorInfo> ParseTemplate<T>(ExcelPackage excelPackage, ImporterAttribute importerAttribute, out List<ImporterHeaderInfo> columnHeaders)
            where T : class
        {
            var tplError = new List<TemplateErrorInfo>();
            //获取导入实体列定义
            ParseImporterHeader<T>(out columnHeaders, out var enumColumns, out var boolColumns);
            try
            {
                //根据名称获取Sheet，如果不存在则取第一个
                var worksheet = excelPackage.Workbook.Worksheets[typeof(T).GetDisplayName()] ?? excelPackage.Workbook.Worksheets[1];
                var excelHeaders = new Dictionary<string, int>();
                for (var i = 0; i < columnHeaders.Count; i++)
                {
                    var header = worksheet.Cells[importerAttribute.HeaderRowIndex, i + 1].Text;
                    //如果出现空格，则截止
                    if (string.IsNullOrWhiteSpace(header))
                        break;

                    if (excelHeaders.ContainsKey(header))
                    {
                        tplError.Add(new TemplateErrorInfo()
                        {
                            ErrorLevel = ErrorLevels.Error,
                            ColumnName = header,
                            RequireColumnName = null,
                            Message = $"列头重复！"
                        });
                    }
                    excelHeaders.Add(header, i + 1);
                }

                foreach (var item in columnHeaders)
                {
                    if (!excelHeaders.ContainsKey(item.ExporterHeader.Name))
                    {
                        //仅验证必填字段
                        if (item.IsRequired)
                        {
                            tplError.Add(new TemplateErrorInfo()
                            {
                                ErrorLevel = ErrorLevels.Error,
                                ColumnName = null,
                                RequireColumnName = item.ExporterHeader.Name,
                                Message = $"当前导入模板中未找到此字段！"
                            });
                            continue;
                        }
                        tplError.Add(new TemplateErrorInfo()
                        {
                            ErrorLevel = ErrorLevels.Warning,
                            ColumnName = null,
                            RequireColumnName = item.ExporterHeader.Name,
                            Message = $"当前导入模板中未找到此字段！"
                        });

                    }
                    else
                    {
                        item.IsExist = true;
                        //设置列索引
                        if (item.ExporterHeader.ColumnIndex == 0)
                            item.ExporterHeader.ColumnIndex = excelHeaders[item.ExporterHeader.Name];
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"模板出现未知错误：{ex.ToString()}");
            }
            return tplError;
        }

       
    }
}