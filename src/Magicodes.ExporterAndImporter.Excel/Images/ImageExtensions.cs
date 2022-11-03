using System.IO;
using System;
using System.Net;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.Metadata;

namespace Magicodes.IE.Excel.Images
{
    internal static class ImageExtensions
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
                var image = Image.Load(wc.OpenRead(url), out format);
                if (image.Metadata.HorizontalResolution == 0 && image.Metadata.VerticalResolution == 0)
                {
                    image.Metadata.HorizontalResolution = ImageMetadata.DefaultHorizontalResolution;
                    image.Metadata.VerticalResolution = ImageMetadata.DefaultVerticalResolution;
                }
                return image;
            }
        }

        /// <summary>
        ///     base64转Bitmap
        /// </summary>
        /// <param name="base64String"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static Image Base64StringToImage(this string base64String, out IImageFormat format)
        {
            var bytes = Convert.FromBase64String(FixBase64ForImage(base64String));
            return Image.Load(bytes, out format);
        }

        private static string FixBase64ForImage(string image)
        {
            var sbText = new System.Text.StringBuilder(image, image.Length);
            sbText.Replace("\r\n", string.Empty);
            sbText.Replace(" ", string.Empty);
            return sbText.ToString();
        }
    }
}
