using Microsoft.IO;
using System;
using System.IO;

namespace OfficeOpenXml.Utils
{
    public static class RecyclableMemoryStream
    {
        private static readonly Lazy<RecyclableMemoryStreamManager> recyclableMemoryStreamManager = new Lazy<RecyclableMemoryStreamManager>();
        private static RecyclableMemoryStreamManager RecyclableMemoryStreamManager
        {
            get
            {
                var recyclableMemoryStream = recyclableMemoryStreamManager.Value;
                recyclableMemoryStream.MaximumFreeSmallPoolBytes = 64 * 1024 * 1024;
                recyclableMemoryStream.MaximumFreeLargePoolBytes = 64 * 1024 * 32;
                recyclableMemoryStream.AggressiveBufferReturn = true;
                return recyclableMemoryStream;
            }
        }
        private const string TagSource = "Magicodes.EPPlus";

        internal static MemoryStream GetStream()
        {
            return RecyclableMemoryStreamManager.GetStream(TagSource);
        }

        internal static MemoryStream GetStream(byte[] array)
        {
            return RecyclableMemoryStreamManager.GetStream(array);
        }

        internal static MemoryStream GetStream(int capacity)
        {
            return RecyclableMemoryStreamManager.GetStream(TagSource, capacity);
        }
    }
}
