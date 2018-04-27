using System;

namespace Magicodes.ExporterAndImporter.Core
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ImporterAttribute : Attribute
    {
        public ImporterAttribute()
        {
        }
    }
}