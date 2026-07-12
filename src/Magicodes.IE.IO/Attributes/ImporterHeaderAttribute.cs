using System;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Specifies the header name used to map a property when importing from a worksheet.
    /// </summary>
    /// <remarks>Apply to a property of the imported model. When <see cref="Name"/> is not set, the property name is matched instead.</remarks>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ImporterHeaderAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the header name to match during import. When omitted, the property name is used.
        /// </summary>
        public string? Name { get; set; }
    }
}
