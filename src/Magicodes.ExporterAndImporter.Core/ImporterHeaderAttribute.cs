using System;

namespace Magicodes.ExporterAndImporter.Core
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ImporterHeaderAttribute : Attribute
    {
        public ImporterHeaderAttribute()
        {
            
        }

        /// <summary>
        /// 列名
        /// </summary>
        public string Name { set; get; }
    }
}