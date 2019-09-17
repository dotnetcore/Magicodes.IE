using System;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// Excel导入表头属性配置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ImporterHeaderAttribute : Attribute
    {
        /// <summary>
        /// 列名
        /// </summary>
        public string Name { set; get; }

        /// <summary>
        /// 批注
        /// </summary>
        public string Description { set; get; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { set; get; } = "X.M";
    }
}