using System.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Bmp;
using SixLabors.ImageSharp.Formats.Gif;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Formats;

namespace Magicodes.IE.EPPlus
{
    public static partial class ImageExtensions
    {
        public static IImageFormat GetImageFormat(this Image image, Stream imageStream)
        {
            var metadataProperty = image.Metadata.GetType().GetProperty("DecodedImageFormat");
            if (metadataProperty != null)
            {
                return metadataProperty.GetValue(image.Metadata) as IImageFormat;
            }
            else
            {
                byte[] header = new byte[4];
                long originalPosition = imageStream.Position;

                imageStream.Position = 0;
                imageStream.Read(header, 0, header.Length);
                imageStream.Position = originalPosition;

                if (header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
                {
                    return PngFormat.Instance;
                }
                if (header[0] == 0xFF && header[1] == 0xD8)
                {
                    return JpegFormat.Instance;
                }
                if (header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46)
                {
                    return GifFormat.Instance;
                }
                if (header[0] == 0x42 && header[1] == 0x4D)
                {
                    return BmpFormat.Instance;
                }
                return null;
            }
        }
    }
}