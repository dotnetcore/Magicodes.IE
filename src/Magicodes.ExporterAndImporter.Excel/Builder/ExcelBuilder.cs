using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Excel.Builder
{
    public class ExcelBuilder
    {
        Func<string, string> ColumnHeaderStringFunc { get; set; }

        private ExcelBuilder() { }

        /// <summary>
        ///     创建实例
        /// </summary>
        /// <returns></returns>
        public static ExcelBuilder Create()
        {
            return new ExcelBuilder();
        }

        /// <summary>
        /// 多语言处理
        /// </summary>
        /// <param name="columnHeaderStringFunc"></param>
        /// <returns></returns>
        public ExcelBuilder WithColumnHeaderStringFunc(Func<string, string>
            columnHeaderStringFunc)
        {
            ColumnHeaderStringFunc = columnHeaderStringFunc;
            return this;
        }

        /// <summary>
        ///     确定设置
        /// </summary>
        public void Build()
        {
            if (ColumnHeaderStringFunc != null)
                ExcelExporter.ColumnHeaderStringFunc = ColumnHeaderStringFunc;
        }

    }
}
