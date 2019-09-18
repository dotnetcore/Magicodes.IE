using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    public class ImportHelper<T> where T : class, new()
    {
        protected ExcelImporterAttribute ExcelImporterAttribute { get; set; }

        protected string FilePath { get; set; }

        protected ImportResult<T> ImportResult { get; set; }

        public ImportHelper(string filePath)
        {
            FilePath = filePath;
        }

        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <returns></returns>
        public Task<ImportResult<T>> Import(string filePath = null)
        {
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                FilePath = filePath;
            }

            ImportResult = new ImportResult<T>();
            try
            {
                var importerAttribute = typeof(T).GetAttribute<ExcelImporterAttribute>(true) ?? new ExcelImporterAttribute();

                CheckImportFile(FilePath);
                ImportResult.RowErrors = new List<DataRowErrorInfo>();
                using (Stream stream = new FileStream(FilePath, FileMode.Open))
                {
                    using (var excelPackage = new ExcelPackage(stream))
                    {
                        #region 检查模板
                        var tplErrors = ParseTemplate<T>(excelPackage, out var columnHeaders);
                        ImportResult.TemplateErrors = tplErrors;
                        if (ImportResult.HasError)
                        {
                            return Task.FromResult(ImportResult);
                        }
                        #endregion

                        ImportResult.Data = ParseData<T>(importerAttribute, excelPackage, columnHeaders);
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
        ///     解析模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private IList<TemplateErrorInfo> ParseTemplate<T>(ExcelPackage excelPackage, out List<ImporterHeaderInfo> columnHeaders)
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

        /// <summary>
        ///     解析数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <exception cref="ArgumentException">最大允许导入条数不能超过5000条</exception>
        private IList<T> ParseData<T>(ImporterAttribute importerAttribute, ExcelPackage excelPackage, List<ImporterHeaderInfo> columnHeaders, IList<DataRowErrorInfo> rowErrors)
            where T : class, new()
        {
            var worksheet = excelPackage.Workbook.Worksheets[typeof(T).GetDisplayName()] ?? excelPackage.Workbook.Worksheets[1];
            if (worksheet.Dimension.End.Row > 5000) throw new ArgumentException("最大允许导入条数不能超过5000条");

            IList<T> importDataModels = new List<T>();
            var propertyInfos = new List<PropertyInfo>(typeof(T).GetProperties());

            for (var rowIndex = importerAttribute.HeaderRowIndex + 1; rowIndex <= worksheet.Dimension.End.Row; rowIndex++)
            {
                int isNullNumber = 1;
                for (int column = 1; column < worksheet.Dimension.End.Column; column++)
                {
                    if (worksheet.Cells[rowIndex, column].Text == string.Empty)
                    {
                        isNullNumber++;
                    }

                }
                if (isNullNumber < worksheet.Dimension.End.Column)
                {
                    var dataItem = new T();
                    foreach (var propertyInfo in propertyInfos)
                    {
                        var col = columnHeaders.Find(a => a.PropertyName == propertyInfo.Name);
                        //检查Excel中是否存在
                        if (!col.IsExist)
                        {
                            continue;
                        }
                        var cell = worksheet.Cells[rowIndex, col.ExporterHeader.ColumnIndex];
                        try
                        {
                            switch (propertyInfo.PropertyType.BaseType?.Name)
                            {
                                case "Enum":
                                    var enumDisplayNames = propertyInfo.PropertyType.GetEnumDisplayNames();
                                    if (enumDisplayNames.ContainsKey(cell.Value?.ToString() ?? throw new ArgumentException()))
                                    {
                                        propertyInfo.SetValue(dataItem,
                                            enumDisplayNames[cell.Value?.ToString()]);
                                    }
                                    else
                                    {
                                        AddRowDataError(rowErrors, rowIndex, col, $"值 {cell.Value} 不存在模板下拉选项中");
                                    }
                                    continue;
                            }
                            var cellValue = cell.Value?.ToString();
                            switch (propertyInfo.PropertyType.GetCSharpTypeName())
                            {
                                case "Boolean":
                                    propertyInfo.SetValue(dataItem, GetBooleanValue(cellValue));
                                    break;
                                case "Nullable<Boolean>":
                                    propertyInfo.SetValue(dataItem, string.IsNullOrWhiteSpace(cellValue) ? (bool?)null : GetBooleanValue(cellValue));
                                    break;
                                case "String":
                                    //TODO:进一步优化
                                    //移除所有的空格，包括中间的空格
                                    if (col.ExporterHeader.FixAllSpace)
                                    {
                                        propertyInfo.SetValue(dataItem, cellValue?.Replace(" ", string.Empty));
                                    }
                                    else if (col.ExporterHeader.AutoTrim)
                                    {
                                        propertyInfo.SetValue(dataItem, cellValue?.Trim());
                                    }
                                    else
                                        propertyInfo.SetValue(dataItem, cellValue?.ToString());
                                    break;
                                //long
                                case "Int64":
                                    {
                                        if (!long.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值，范围为{long.MinValue}~{long.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int64>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        if (!long.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值，范围为{long.MinValue}~{long.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Int32":
                                    {
                                        if (!int.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值，范围为{int.MinValue}~{int.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Int32>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        if (!int.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的整数数值，范围为{int.MinValue}~{int.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Decimal":
                                    {
                                        if (!decimal.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数，范围为{decimal.MinValue}~{decimal.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Decimal>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        if (!decimal.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数，范围为{decimal.MinValue}~{decimal.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Double":
                                    {
                                        if (!double.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数，范围为{double.MinValue}~{double.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Double>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        if (!double.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数，范围为{double.MinValue}~{double.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                //case "float":
                                case "Single":
                                    {
                                        if (!float.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数，范围为{float.MinValue}~{float.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "Nullable<Single>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        if (!float.TryParse(cellValue, out var number))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的小数，范围为{float.MinValue}~{float.MaxValue}！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, number);
                                    }
                                    break;
                                case "DateTime":
                                    {
                                        if (!DateTime.TryParse(cellValue, out var date))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的日期时间格式！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, date);
                                    }
                                    break;
                                case "Nullable<DateTime>":
                                    {
                                        if (string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            propertyInfo.SetValue(dataItem, null);
                                            break;
                                        }
                                        if (!DateTime.TryParse(cellValue, out var date))
                                        {
                                            AddRowDataError(rowErrors, rowIndex, col, $"值 {cellValue} 无效，请填写正确的日期时间格式！");
                                            break;
                                        }
                                        propertyInfo.SetValue(dataItem, date);
                                    }
                                    break;
                                default:
                                    propertyInfo.SetValue(dataItem, cell.Value);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            AddRowDataError(rowErrors, rowIndex, col, ex.Message);
                        }
                    }

                    importDataModels.Add(dataItem);
                }
            }

            return importDataModels;
        }

        /// <summary>
        /// 添加数据行错误
        /// </summary>
        /// <param name="rowErrors"></param>
        /// <param name="rowIndex"></param>
        /// <param name="importerHeaderInfo"></param>
        private void AddRowDataError(IList<DataRowErrorInfo> rowErrors, int rowIndex, ImporterHeaderInfo importerHeaderInfo, string errorMessage = "数据格式无效！")
        {
            var rowError = rowErrors.FirstOrDefault(p => p.RowIndex == rowIndex);
            if (rowError == null)
            {
                rowError = new DataRowErrorInfo()
                {
                    RowIndex = rowIndex
                };
                rowErrors.Add(rowError);
            }
            rowError.FieldErrors.Add(importerHeaderInfo.ExporterHeader.Name, errorMessage);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static bool GetBooleanValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            switch (value.ToLower())
            {
                case "1":
                case "是":
                case "yes":
                case "true":
                    return true;
                case "0":
                case "否":
                case "no":
                case "false":
                default:
                    return false;
            }
        }
    }
}
