using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Excel.Builder
{
    public class ExcelBuilder
    {
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
        ///     设置日志记录处理
        /// </summary>
        /// <param name="logger"></param>
        /// <returns></returns>
        public ExcelBuilder WithLoggerAction(Action<string, string> loggerAction)
        {
            //LoggerAction = loggerAction;
            return this;
        }

        /// <summary>
        ///     确定设置
        /// </summary>
        public void Build()
        {
            
        }

    }
}
