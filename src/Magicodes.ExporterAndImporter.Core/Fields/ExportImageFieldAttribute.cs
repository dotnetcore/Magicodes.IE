using System;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    ///     导出图片字段特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportImageFieldAttribute : Attribute
    {
        /// <summary>
        ///     高度
        /// </summary>
        //TODO 高度始终会按照最后一列高度来设置,考虑是否可迁移
        public int Height { get; set; } = 15;
        /// <summary>
        ///     宽度
        /// </summary>
        public int Width { get; set; } = 50;
        /// <summary>
        ///     图片不存在时的替代文本
        /// </summary>
        public string Alt { get; set; }

        /// <summary>
        /// 垂直偏移
        /// </summary>
        public int YOffset { get; set; } = 0;

        /// <summary>
        /// 水平偏移
        /// </summary>
        public int XOffset { get; set; } = 0;

        /// <summary>
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="alt"></param>
        public ExportImageFieldAttribute(int height = 15, int width = 50, string alt = null)
        {
            Height = height;
            Width = width;
            Alt = alt;
        }
    }
}
