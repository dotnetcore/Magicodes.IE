using System;

namespace Magicodes.ExcelImporter.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ImporterAttribute : Attribute
    {
        public ImporterAttribute()
        {
            
        }

        /// <summary>
        /// 字体大小
        /// </summary>
        public float? FontSize { set; get; }
    }
}