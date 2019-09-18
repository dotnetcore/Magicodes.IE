using System;

namespace Magicodes.ExporterAndImporter.Core
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ImporterAttribute : Attribute
    {
        public ImporterAttribute()
        {
        }

        /// <summary>
        /// 表头位置
        /// </summary>
        public int HeaderRowIndex { get; set; } = 1;
    }
}