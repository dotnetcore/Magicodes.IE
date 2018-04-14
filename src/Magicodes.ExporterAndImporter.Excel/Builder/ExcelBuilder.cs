using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Excel.Builder
{
    public class ExcelBuilder
    {
        Func<string, string> LocalStringFunc { get; set; }

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
        /// <param name="localStringFunc"></param>
        /// <returns></returns>
        public ExcelBuilder WithLocalStringFunc(Func<string, string> localStringFunc)
        {
            LocalStringFunc = localStringFunc;
            return this;
        }

        /// <summary>
        ///     确定设置
        /// </summary>
        public void Build()
        {
            if (LocalStringFunc != null)
                ExcelExporter.LocalStringFunc = LocalStringFunc;
        }

    }
}
