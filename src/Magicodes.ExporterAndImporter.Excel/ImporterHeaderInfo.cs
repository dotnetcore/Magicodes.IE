// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : ImporterHeaderInfo.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     导入列头设置
    /// </summary>
    public class ImporterHeaderInfo
    {
        /// <summary>
        ///     是否必填
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        ///     列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        ///     列属性
        /// </summary>
        public ImporterHeaderAttribute ExporterHeader { get; set; }

        /// <summary>
        ///     是否存在
        /// </summary>
        internal bool IsExist { get; set; }
    }
}