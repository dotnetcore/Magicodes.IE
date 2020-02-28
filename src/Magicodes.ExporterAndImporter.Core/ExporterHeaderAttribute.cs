// ======================================================================
// 
//           filename : ExporterHeaderAttribute.cs
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
    [AttributeUsage(AttributeTargets.Property)]
    public class ExporterHeaderAttribute : Attribute
    {
        public ExporterHeaderAttribute(string displayName = null, float fontSize = 11, string format = null,
            bool isBold = true, bool isAutoFit = true,bool isImg=false,int imgHeight=0,int imgWidth=0,string imgIsNullText="")
        {
            DisplayName = displayName;
            FontSize = fontSize;
            Format = format;
            IsBold = isBold;
            IsAutoFit = isAutoFit;
            IsImg = isImg;
            ImgHeight = imgHeight;
            ImgWidth = imgWidth;
            ImgIsNullText = imgIsNullText;
        }

        /// <summary>
        ///     显示名称
        /// </summary>
        public string DisplayName { set; get; }

        /// <summary>
        ///     字体大小
        /// </summary>
        public float? FontSize { set; get; }

        /// <summary>
        ///     是否加粗
        /// </summary>
        public bool IsBold { set; get; }

        /// <summary>
        ///     格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        ///     是否自适应
        /// </summary>
        public bool IsAutoFit { set; get; }

        /// <summary>
        ///     是否忽略
        /// </summary>
        public bool IsIgnore { get; set; }
        /// <summary>
        ///     是否是图片
        /// </summary>
        public bool IsImg { get; set; }
        /// <summary>
        ///     高度
        /// </summary>
        //TODO 高度始终会按照最后一列高度来设置,考虑是否可迁移
        public int ImgHeight { get; set; }
        /// <summary>
        ///     宽度
        /// </summary>
        public int ImgWidth { get; set; }
        /// <summary>
        ///     图片不存在默认填充数据
        /// </summary>
        public string ImgIsNullText { get; set; }
    }



}