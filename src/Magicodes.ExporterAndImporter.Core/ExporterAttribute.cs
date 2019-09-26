// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : ExporterAttribute.cs
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

namespace Magicodes.ExporterAndImporter.Core
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExporterAttribute : Attribute
    {
        /// <summary>
        ///     名称(比如当前Sheet 名称)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        ///     头部字体大小
        /// </summary>
        public float? HeaderFontSize { set; get; }

        /// <summary>
        ///     正文字体大小
        /// </summary>
        public float? FontSize { set; get; }
    }
}