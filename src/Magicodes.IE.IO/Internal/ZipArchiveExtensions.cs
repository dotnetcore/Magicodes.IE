
using System.IO;
using System.IO.Compression;

namespace Magicodes.IE.IO
{
    internal static class ZipArchiveExtensions
    {
        internal static Stream OpenEntry(this ZipArchive archive, string name, CompressionLevel compression)
        {
            var entry = archive.CreateEntry(name, compression);
            return entry.Open();
        }
    }
}
