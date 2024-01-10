using System.IO;
using System;
using System.Net;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.Metadata;
using SkiaSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Advanced;
using SixLabors.ImageSharp.Metadata.Profiles.Exif;
using Magicodes.IE.EPPlus;
using SixLabors.ImageSharp.Memory;
using System.Text.RegularExpressions;

namespace Magicodes.IE.Excel.Images
{
    internal static partial class ImageExtensions
    {
        public static string SaveTo(this Image image, string path)
        {
            image.Save(path);
            return path;
        }

        public static string ToBase64String(this Image image, IImageFormat format)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, format);
                ms.Position = 0;
                var bytes = ms.ToArray();
                return Convert.ToBase64String(bytes);
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="url"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static Image GetImageByUrl(this string url, out IImageFormat format)
        {
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            using (var wc = new WebClient())
            {
                wc.Proxy = null;
                using (Stream webStream = wc.OpenRead(url))
                {
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        webStream.CopyTo(memoryStream);
                        memoryStream.Position = 0;

                        var image = Image.Load(memoryStream);
                        format = image.GetImageFormat(memoryStream);

                        if (image.Metadata.HorizontalResolution == 0 && image.Metadata.VerticalResolution == 0)
                        {
                            image.Metadata.HorizontalResolution = ImageMetadata.DefaultHorizontalResolution;
                            image.Metadata.VerticalResolution = ImageMetadata.DefaultVerticalResolution;
                        }
                        return image;
                    }
                }
            }
        }

        /// <summary>
        ///     Converts a base64 string to an Image
        /// </summary>
        /// <param name="base64String">The base64 string representing the image</param>
        /// <param name="format">The image format</param>
        /// <returns>An Image object representing the base64 string</returns>
        public static Image Base64StringToImage(this string base64String, out IImageFormat format)
        {
            byte[] bytes = Convert.FromBase64String(CleanupBase64String(base64String));
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                Image image = Image.Load(stream);
                format = image.GetImageFormat(stream);
                return image;
            }
        }

        /// <summary>
        ///     Cleans up the base64 string by removing unnecessary characters
        /// </summary>
        /// <param name="base64String">The base64 string to clean up</param>
        /// <returns>A cleaned-up base64 string</returns>
        private static string CleanupBase64String(string base64String)
        {
            return Regex.Replace(base64String, @"\s+", string.Empty);
        }
    }
}
