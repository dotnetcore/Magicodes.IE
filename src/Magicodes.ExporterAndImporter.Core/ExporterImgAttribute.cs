using System;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    ///     导出图片特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExporterImgAttribute : Attribute
    {
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
        /// <summary>
        /// </summary>
        /// <param name="isImg"></param>
        /// <param name="imgHeight"></param>
        /// <param name="imgWidth"></param>
        /// <param name="imgIsNullText"></param>
        public ExporterImgAttribute(bool isImg, int imgHeight = 15, int imgWidth = 50, string imgIsNullText = null)
        {
            IsImg = isImg;
            ImgHeight = imgHeight;
            ImgWidth = imgWidth;
            ImgIsNullText = imgIsNullText;
        }
        /// <summary>
        /// </summary>
        /// <param name="isImg"></param>
        public ExporterImgAttribute(bool isImg)
        {
            IsImg = isImg;
            ImgHeight = 15;
            ImgWidth = 50;
            ImgIsNullText = null;
        }
    }
}
