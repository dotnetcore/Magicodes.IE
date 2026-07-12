using System;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Enables source-generated metadata so the annotated type can be read from and written to .xlsx without reflection.
    /// </summary>
    /// <remarks>Apply to a class or struct that you export or import. The generated reader and writer avoid reflection and reduce allocations.</remarks>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, Inherited = false)]
    public sealed class XlsxExportableAttribute : Attribute
    {
    }
}
