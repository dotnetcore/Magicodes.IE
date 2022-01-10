using System;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    ///     忽略特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class IEIgnoreAttribute : Attribute
    {
        /// <summary>
        ///     是否忽略导入，默认true
        /// </summary>
        public bool IsImportIgnore { get; set; } = true;

        /// <summary>
        ///     是否忽略导出，默认true
        /// </summary>
        public bool IsExportIgnore { get; set; } = true;
    }
}