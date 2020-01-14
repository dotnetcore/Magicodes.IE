// ======================================================================
// 
//           filename : TemplateExportHelper.cs
//           description :
// 
//           created by 雪雁 at  2020-01-06 16:13
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DynamicExpresso;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using OfficeOpenXml;

namespace Magicodes.ExporterAndImporter.Excel.Utility.TemplateExport
{
    /// <summary>
    ///     模板导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class TemplateExportHelper<T> : IDisposable where T : class
    {
        /// <summary>
        ///     匹配表达式
        /// </summary>
        private const string VariableRegexString = "(\\{\\{)+([\\w_.>|]*)+(\\}\\})";

        private ExcelExporterAttribute _excelExporterAttribute;

        /// <summary>
        ///     变量正则
        /// </summary>
        private readonly Regex _variableRegex = new Regex(VariableRegexString, RegexOptions.IgnoreCase);

        /// <summary>
        ///     模板文件路径
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        ///     模板文件路径
        /// </summary>
        public string TemplateFilePath { get; set; }

        /// <summary>
        ///     模板写入器
        /// </summary>
        protected Dictionary<string, List<IWriter>> SheetWriters { get; set; }

        /// <summary>
        ///     模板文件导出结果
        /// </summary>
        protected TemplateExportResult TemplateExportResult { get; set; }

        /// <summary>
        ///     值字典
        /// </summary>
        protected Dictionary<string, string> ValuesDictionary { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        protected T Data { get; set; }

        /// <summary>
        ///     表值字典
        /// </summary>
        protected Dictionary<string, List<Dictionary<string, string>>> TableValuesDictionary { get; set; }

        ///// <summary>
        /////     导出全局设置
        ///// </summary>
        //protected ExcelExporterAttribute ExcelExporterSettings
        //{
        //    get
        //    {
        //        if (_excelExporterAttribute == null)
        //        {
        //            var type = typeof(T);
        //            _excelExporterAttribute = type.GetAttribute<ExcelExporterAttribute>(true);
        //            if (_excelExporterAttribute != null) return _excelExporterAttribute;

        //            var importerAttribute = type.GetAttribute<ExporterAttribute>(true);
        //            if (importerAttribute != null)
        //                _excelExporterAttribute = new ExcelExporterAttribute();
        //            else
        //                _excelExporterAttribute = new ExcelExporterAttribute();

        //            return _excelExporterAttribute;
        //        }

        //        return _excelExporterAttribute;
        //    }
        //    set => _excelExporterAttribute = value;
        //}




        /// <summary>
        ///     根据模板导出Excel
        /// </summary>
        /// <param name="filePath">导出文件路径</param>
        /// <param name="templateFilePath">模板文件路径</param>
        /// <param name="data"></param>
        public void Export(string filePath, string templateFilePath, T data)
        {
            if (!string.IsNullOrWhiteSpace(filePath)) FilePath = filePath;
            if (!string.IsNullOrWhiteSpace(templateFilePath)) TemplateFilePath = templateFilePath;

            if (string.IsNullOrWhiteSpace(FilePath)) throw new ArgumentException("文件路径不能为空!", nameof(FilePath));
            if (string.IsNullOrWhiteSpace(TemplateFilePath))
                throw new ArgumentException("模板文件路径不能为空!", nameof(TemplateFilePath));

            if (data == null) throw new ArgumentException("数据不能为空!", nameof(data));

            Data = data;

            using (Stream stream = new FileStream(TemplateFilePath, FileMode.Open))
            {
                using (var excelPackage = new ExcelPackage(stream))
                {
                    ParseTemplateFile(excelPackage);

                    ParseData(excelPackage);

                    excelPackage.SaveAs(new FileInfo(FilePath));
                }
            }
        }

        /// <summary>
        /// 处理数据
        /// </summary>
        /// <param name="excelPackage"></param>
        private void ParseData(ExcelPackage excelPackage)
        {
            var target = new Interpreter();
            target.SetVariable("data", Data, typeof(T));
            
            //表格渲染参数
            var tbParameters = new[] {
                //new Parameter("data", typeof(T)),
                new Parameter("index", typeof(int))
            };

            //TODO:渲染支持自定义处理程序
            foreach (var sheetName in SheetWriters.Keys)
            {
                var sheet = excelPackage.Workbook.Worksheets[sheetName];

                #region 处理普通单元格模板
                foreach (var writer in SheetWriters[sheetName].Where(p => p.WriterType == WriterTypes.Cell))
                {
                    var expresson = writer.CellString
                        .Replace("{{", "\" + data.")
                        .Replace("}}", " + \"");

                    expresson = expresson.StartsWith("\"")
                        ? expresson.TrimStart('\"').TrimStart().TrimStart('+')
                        : "\"" + expresson;

                    expresson = expresson.EndsWith("\"")
                        ? expresson.TrimEnd('\"').TrimEnd().TrimEnd('+')
                        : expresson + "\"";

                    var cellWriteFunc = target.Parse(expresson);
                    var result = cellWriteFunc.Invoke();
                    sheet.Cells[writer.Address].Value = result;
                }

                #endregion

                #region 处理表格

                var tableGroups = SheetWriters[sheetName].Where(p => p.WriterType == WriterTypes.Table)
                    .GroupBy(p => p.TableKey);


                foreach (var tableGroup in tableGroups)
                {
                    var tableKey = tableGroup.Key;
                    //TODO:处理异常“No property or field”
                    var rowCount = target.Eval<int>($"data.{tableKey}.Count");
                    if (rowCount == 0) return;

                    Console.WriteLine($"正在处理表格【{tableKey}】，行数：{rowCount}。");
                    var isFirst = true;
                    foreach (var item in tableGroup)
                    {
                        var address = new ExcelAddressBase(item.Address);
                        //TODO:支持同一行多个表格
                        if (isFirst && rowCount > 1)
                        {
                            //插入行
                            sheet.InsertRow(address.Start.Row + 1, rowCount - 1, address.Start.Row);
                        }
                        isFirst = false;

                        var cellString = item.CellString;
                        if (cellString.Contains("{{Table>>"))
                            //{{ Table >> BookInfo | RowNo}}
                            cellString = "{{" + cellString.Split('|')[1].Trim();
                        else if (cellString.Contains(">>Table}}"))
                            //{{Remark|>>Table}}
                            cellString = cellString.Split('|')[0].Trim() + "}}";

                        var expresson = cellString
                            .Replace("{{", "\" + data." + tableKey + "[index].")
                            .Replace("}}", " + \"");

                        expresson = expresson.StartsWith("\"")
                            ? expresson.TrimStart('\"').TrimStart().TrimStart('+')
                            : "\"" + expresson;

                        expresson = expresson.EndsWith("\"")
                            ? expresson.TrimEnd('\"').TrimEnd().TrimEnd('+')
                            : expresson + "\"";

                        var cellWriteFunc = target.Parse(expresson, tbParameters);

                        for (var i = 0; i < rowCount; i++)
                        {
                            var result = cellWriteFunc.Invoke(i);
                            sheet.Cells[address.Start.Row + i, address.Start.Column].Value = result;
                        }
                    }
                }

                #endregion
            }
        }

        /// <summary>
        ///     验证并转换模板
        /// </summary>
        /// <param name="excelPackage"></param>
        protected void ParseTemplateFile(ExcelPackage excelPackage)
        {
            SheetWriters = new Dictionary<string, List<IWriter>>();
            foreach (var worksheet in excelPackage.Workbook.Worksheets)
            {
                var writers = new List<IWriter>();
                if (!SheetWriters.ContainsKey(worksheet.Name)) SheetWriters.Add(worksheet.Name, writers);
                var endColumnIndex = worksheet.Dimension.End.Column;
                var endRowIndex = worksheet.Dimension.End.Row;

                //获取所有包含表达式的单元格
                var q = (from cell in worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column,
                        endRowIndex, endColumnIndex]
                         where _variableRegex.IsMatch((cell.Value ?? string.Empty).ToString())
                         select cell).ToList();

                var rows = q.GroupBy(p => p.Rows);


                foreach (var rowGroups in rows)
                {
                    var isStartTable = false;
                    string tableKey = null;
                    foreach (var cell in rowGroups)
                    {
                        var cellString = cell.Value.ToString();
                        if (cellString.Contains("{{Table>>"))
                        {
                            isStartTable = true;
                            //{{ Table >> BookInfo | RowNo}}
                            tableKey = Regex.Split(cellString, "{{Table>>")[1].Split('|')[0].Trim();
                        }

                        writers.Add(new Writer
                        {
                            TableKey = tableKey,
                            Address = cell.Address,
                            CellString = cellString,
                            WriterType = isStartTable ? WriterTypes.Table : WriterTypes.Cell,
                        });

                        if (isStartTable && cellString.Contains(">>Table}}"))
                        {
                            isStartTable = false;
                            tableKey = null;
                        }
                    }
                }
            }
        }

        /// <summary>执行与释放或重置非托管资源关联的应用程序定义的任务。</summary>
        public void Dispose()
        {
            FilePath = null;
            SheetWriters = null;
            ValuesDictionary = null;
            TableValuesDictionary = null;
        }
    }
}