using System;

namespace Magicodes.ExporterAndImporter.Excel.Builder
{
    /// <summary>
    /// Excel表头多语言处理
    /// </summary>
    public class ExcelBuilder
    {
        private Func<string, string> ColumnHeaderStringFunc { get; set; }

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
            {
                ExcelExporter.ColumnHeaderStringFunc = ColumnHeaderStringFunc;
            }
        }
    }
}