// ======================================================================
// 
//           filename : ImportResult.cs
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
using System.Collections.Generic;
using System.Linq;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    ///     导入结果
    /// </summary>
    public class ImportResult<T> where T : class
    {
        /// <summary>
        /// </summary>
        public ImportResult()
        {
            RowErrors = new List<DataRowErrorInfo>();
        }

        /// <summary>
        ///     导入数据
        /// </summary>
        public virtual ICollection<T> Data { get; set; }

        /// <summary>
        ///     验证错误
        /// </summary>
        public virtual IList<DataRowErrorInfo> RowErrors { get; set; }

        /// <summary>
        ///     模板错误
        /// </summary>
        public virtual IList<TemplateErrorInfo> TemplateErrors { get; set; }

        /// <summary>
        ///     导入异常信息
        /// </summary>
        public virtual Exception Exception { get; set; }

        /// <summary>
        ///     是否存在导入错误
        /// </summary>
        public virtual bool HasError => Exception != null ||
                                        (TemplateErrors?.Count(p => p.ErrorLevel == ErrorLevels.Error) ?? 0) > 0 ||
                                        (RowErrors?.Count ?? 0) > 0;

        /// <summary>
        ///     Imported header list information
        ///     导入的表头列表信息
        ///     https://github.com/dotnetcore/Magicodes.IE/issues/76
        /// </summary>
        public virtual IList<ImporterHeaderInfo> ImporterHeaderInfos { get; set; }
    }
}