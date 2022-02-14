using OfficeOpenXml.Utils;
using System.Drawing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats;

namespace OfficeOpenXml.Compatibility
{
    internal class ImageCompat
    {
        internal static byte[] GetImageAsByteArray(Image image)
        {
            var ms = RecyclableMemoryStream.GetStream();
            //IImageFormat format = GetFormat(extension);

            if (image.RawFormat.Guid == ImageFormat.Gif.Guid)
            {
                image.Save(ms, ImageFormat.Gif);
            }
            else if (image.RawFormat.Guid == ImageFormat.Bmp.Guid)
            {
                image.Save(ms, ImageFormat.Bmp);
            }
            else if (image.RawFormat.Guid == ImageFormat.Png.Guid)
            {
                image.Save(ms, ImageFormat.Png);
            }
            else if (image.RawFormat.Guid == ImageFormat.Tiff.Guid)
            {
                image.Save(ms, ImageFormat.Tiff);
            }
            else
            {
                image.Save(ms, ImageFormat.Jpeg);
            }

            return ms.ToArray();
        }
    }
}
