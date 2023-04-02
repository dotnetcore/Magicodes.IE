using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CSScriptLib;
using System.Text;
using CSScripting;
using System.Dynamic;
using Magicodes.IE.Stash.Extensions;

namespace Magicodes.IE.Stash.Import
{
    public partial class ExcelImporter
    {
        /// <summary>
        /// 映射定义
        /// </summary>
        public ImportMapDefinition? MapDefinition { get; set; }

        /// <summary>
        /// 从json字符串中加载映射规则
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public ImportMapDefinition LoadDefinitionFromJson(string json)
        {
            //TODO: 这里应该由用户决定Json序列化器.
            throw new NotImplementedException();
        }

        /// <summary>
        /// 从excel文件中加载映射规则
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public ImportMapDefinition LoadDefinitionFromExcelFile(string excelFilePath)
        {
            var ret = new ImportMapDefinition();

            var fi = new FileInfo(excelFilePath);
            if (!fi.Exists)
            {
                throw new FileNotFoundException("指定的文件不存在", excelFilePath);
            }
            using (var package = new ExcelPackage(fi))
            {
                var workBook = package.Workbook;
                var workSheets = workBook.Worksheets;
                var sheet = workSheets[0];
                var rowsCount = sheet.Dimension.Rows;
                var colsCount = sheet.Dimension.Columns;

                var titleRowIdx = -1;
                var titles = new List<string>() { null };

                for (int rowIdx = 1; rowIdx <= rowsCount; rowIdx++)
                {
                    var posCol = sheet.Cells[rowIdx, 1];
                    var posText = posCol.Text;
                    if (posText == "^定义完^")
                    {
                        break;
                    }
                    if (titleRowIdx > 0)//已经过了标题行
                    {

                        var mapItem = new MapItem() { };
                        ret.Maps.Add(mapItem);

                        for (int i = 1; i < titles.Count; i++)
                        {
                            var value = sheet.Cells[rowIdx, i].Text;
                            switch (titles[i])
                            {
                                case "序号":
                                    mapItem.Index = value;
                                    break;
                                case "数据源列":
                                    mapItem.Column = value;
                                    break;
                                case "模型属性名":
                                    //TODO: 还没搞子属性
                                    mapItem.Property = value;
                                    break;
                                case "数据类型":
                                    mapItem.Type = value;
                                    break;
                                case "默认值":
                                    mapItem.Default = value;
                                    break;
                                case "异常处理方式":
                                    mapItem.Fail = value;
                                    break;
                                case "转换器":
                                    if (!string.IsNullOrWhiteSpace(value))
                                    {
                                        mapItem.Pipes.Add(new() { Code = value });
                                    }
                                    break;
                            }
                        }

                    }
                    else
                    {
                        switch (posText)
                        {
                            //读取Dto类型
                            case "模型类型:":
                                ret.DtoTypeName = sheet.Cells[rowIdx, posCol.Columns + 1].Text;
                                break;
                            case "命名空间:":
                                var txt = sheet.Cells[rowIdx, posCol.Columns + 1].Text;
                                var sp = txt.Split(new char[] { ';' });
                                foreach (var item in sp)
                                {
                                    var v = item.Trim();
                                    if (!item.Equals("using", StringComparison.OrdinalIgnoreCase))
                                    {
                                        ret.Namespaces.Add(v);
                                    }
                                }
                                break;
                            //读取变量定义
                            case "变量:":
                                ret.Variables.Add(new()
                                {
                                    Name = sheet.Cells[rowIdx, posCol.Columns + 1].Text,
                                    Code = sheet.Cells[rowIdx, posCol.Columns + 2].Text
                                });

                                break;
                            //读取映射定义
                            case "序号":
                                if (titleRowIdx == -1)
                                {
                                    //发现 "序号"二字,记下它的行号,这行以后的所有行,都看作是映射项,只到发现映射结束字符,或所有行都搞完
                                    titleRowIdx = rowIdx;
                                    for (int colIdx = 1; colIdx <= colsCount; colIdx++)
                                    {
                                        var title = sheet.Cells[rowIdx, colIdx].Text;
                                        titles.Add(title);
                                    }
                                }
                                break;
                        }
                    }
                }
            }

            MapDefinition = ret;
            return ret;
        }

        /// <summary>
        /// 解析数据
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>

        //TODO: 方法是否需要做成异步的?
        public List<object> Resolve(string filePath)
        {
            var _def = MapDefinition;
            var type = FindType(_def.DtoTypeName);
            if (type == null)
            {
                throw new Exception($"未找到指定的Dto类型:{_def.DtoTypeName} , 请尝试提供带命名空间的完整类型名");
            }
            var _变量计算期context = new ImportContext()
            {
                Def = _def,
                DtoType = type,
                Variables = new ExpandoObject()
            };


            foreach (var item in _def.Variables)
            {
                var val = item.Func.Calc(_变量计算期context);
                item.Value = val;

                ((IDictionary<string, object?>)_变量计算期context.Variables).Add(item.Name, val);
            }

            var ret = new List<object>();

            /// 目标类型的属性定义
            //TODO: 还没搞子属性
            var _propertieInfos = type.GetProperties()
                .Where(p => p.IsPublic() && p.CanRead && p.CanBeSet())
                .ToList();

            var fi = new FileInfo(filePath);
            using (var package = new ExcelPackage(fi))
            {
                var workBook = package.Workbook;
                var workSheets = workBook.Worksheets;
                var sheet = workSheets[0];
                var rowsCount = sheet.Dimension.Rows;
                var colsCount = sheet.Dimension.Columns;

                var _titles = new List<string>() { null };//加一个null,是因为epplus的索引是从1开始的,为了方便而已

                //TODO: 标题行所在的行号,这里应该搞成可配置的,因为有些场景,用户excel文件的标题可能不在第一行.
                var titleRowIdx = 1;

                for (int colIdx = 1; colIdx <= colsCount; colIdx++)
                {
                    var _col = sheet.Cells[titleRowIdx, colIdx];
                    _titles.Add(_col.Text);
                }

                for (int rowIdx = titleRowIdx + 1; rowIdx <= rowsCount; rowIdx++)
                {
                    List<CellValue> colValues = new();
                    for (int colIdx = 1; colIdx <= colsCount; colIdx++)
                    {
                        var cell = sheet.Cells[rowIdx, colIdx];
                        colValues.Add(new()
                        {
                            Index = colIdx,
                            ExcelRange = cell,
                            Text = cell.Text,
                            Title = _titles[colIdx],
                            Value = cell.Value,
                        });
                    }

                    var _dtoInstance = Activator.CreateInstance(type)!;
                    ret.Add(_dtoInstance);

                    foreach (var _map in _def.Maps)
                    {
                        if (string.IsNullOrWhiteSpace(_map.Property))
                        {
                            //TODO: 如果属性为空,是否应该处理后续的转换器?,思考有二:
                            // 如果转换器仅仅只是计算返回值,那没必要
                            // 但转换器里注入的对象,是允许修改的,导致转换器可以直接修改对象(dto实例其它属性)的值,会产生影响,同时,也提供了灵活性.
                            continue;
                        }

                        var prInfo = _propertieInfos.SingleOrDefault(p => p.Name == _map.Property);

                        if (prInfo == null)
                        {
                            throw new Exception($"类型{type.Name}不存在指定的属性{_map.Property}.");
                        }

                        object value = null;

                        if (!string.IsNullOrWhiteSpace(_map.Column))
                        {
                            var _titleIdx = _titles.IndexOf(_map.Column);

                            if (_titleIdx > 0)
                            {
                                value = sheet.Cells[rowIdx, _titleIdx].GetCellValueByType(prInfo.PropertyType);
                            }
                            else if (_map.Column.Trim().StartsWith("{{") && _map.Column.Trim().EndsWith("}}"))
                            {
                                var name = _map.Column.Trim().TrimStart('{').TrimEnd('}').Trim();
                                switch (name)
                                {
                                    case "全部":
                                        value = colValues;
                                        break;
                                    case "其它":
                                        break;
                                    default:
                                        value = sheet.Cells[$"{name}{rowIdx}"].GetCellValueByType(prInfo.PropertyType);
                                        break;
                                }
                            }
                            else
                            {
                                // 数据源中没找到指定的列,用默认值代替
                                if (_map.Default != null)
                                {
                                    value = _map.Default.GetCellValueByType(prInfo.PropertyType);
                                }
                            }
                        }


                        foreach (var pipe in _map.Pipes)
                        {
                            var context = new ImportMapContext()
                            {
                                DtoObj = _dtoInstance,
                                Map = _map,
                                Code = pipe.Code,
                                Def = _def,
                                Value = value,
                                Cells = colValues,
                                Variables = _变量计算期context.Variables,
                            };

                            value = pipe.Func.Calc(context);
                        }

                        prInfo.SetValue(_dtoInstance, value);
                    }
                }
            }

            return ret;

        }
    }
}
