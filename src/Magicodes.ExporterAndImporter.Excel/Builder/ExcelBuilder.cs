// ======================================================================
// 
//           filename : ExcelBuilder.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;

namespace Magicodes.ExporterAndImporter.Excel.Builder
{
    public class ExcelBuilder
    {
        private ExcelBuilder()
        {
        }

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
        ///     多语言处理
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