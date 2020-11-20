// ======================================================================
// 
//           filename : ImportRowDataErrorDto.cs
//           description :
// 
//           created by 雪雁 at  2019-11-05 20:02
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System.Collections.Generic;
using System.Linq;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public class ImportResultFilterTest : IImportResultFilter
    {
        /// <summary>
        /// 本示例修改数据错误验证结果，可用于多语言等场景
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="importResult"></param>
        /// <returns></returns>
        public ImportResult<T> Filter<T>(ImportResult<T> importResult) where T : class, new()
        {
            var errorRows = new List<int>()
            {
                5,6
            };
            var items = importResult.RowErrors.Where(p => errorRows.Contains(p.RowIndex)).ToList();

            for (int i = 0; i < items.Count; i++)
            {
                for (int j = 0; j < items[i].FieldErrors.Keys.Count; j++)
                {
                    var key = items[i].FieldErrors.Keys.ElementAt(j);
                    var value = items[i].FieldErrors[key];
                    items[i].FieldErrors[key] = value?.Replace("存在数据重复，请检查！所在行：", "Duplicate data exists, please check! Where:");
                }
            }

            foreach (var item in importResult.Data)
            {
                //判断两个列是否重复，如果重复则提示 https://github.com/dotnetcore/Magicodes.IE/issues/144
                if (item is ImportResultFilterDataDto1 item1 && item1.Code == item1.Name)
                {
                    items.Add(new DataRowErrorInfo()
                    {
                        RowIndex = 2, //请按照正确的行号去填写
                        FieldErrors = new Dictionary<string, string>()
                        {
                            {"产品名称","Code和Name重复"}
                        }
                    });
                }
            }

            importResult.RowErrors = items;

            return importResult;
        }
    }

    [ExcelImporter(ImportResultFilter = typeof(ImportResultFilterTest), IsLabelingError = true)]
    public class ImportResultFilterDataDto1
    {
        /// <summary>
        ///     产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称")]
        public string Name { get; set; }

        /// <summary>
        ///     产品代码
        ///     长度验证
        ///     重复验证
        /// </summary>
        [ImporterHeader(Name = "产品代码", Description = "最大长度为20", AutoTrim = false, IsAllowRepeat = false)]
        public string Code { get; set; }
    }

    [ExcelImporter(IsLabelingError = true)]
    public class DIImportResultFilterDataDto1
    {
        /// <summary>
        ///     产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称")]
        public string Name { get; set; }

        /// <summary>
        ///     产品代码
        ///     长度验证
        ///     重复验证
        /// </summary>
        [ImporterHeader(Name = "产品代码", Description = "最大长度为20", AutoTrim = false, IsAllowRepeat = false)]
        public string Code { get; set; }
    }
}