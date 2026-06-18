using System.IO;
using System;
using System.Net.Http;
using SkiaSharp;
using SixLabors.ImageSharp;
using Magicodes.IE.EPPlus;

namespace Magicodes.IE.Excel.Images
{
    public static partial class ImageExtensions
    {
        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

        public struct ImageInfo
        {
            public int Width;
            public int Height;
            public string ContentType;
        }

        public static ImageInfo IdentifyImage(byte[] imageBytes)
        {
            using (var ms = new MemoryStream(imageBytes))
            {
                using (var codec = SKCodec.Create(ms))
                {
                    if (codec == null)
                        throw new InvalidOperationException("无法识别图片格式");

                    return new ImageInfo
                    {
                        Width = codec.Info.Width,
                        Height = codec.Info.Height,
                        ContentType = GetContentType(codec.EncodedFormat)
                    };
                }
            }
        }

        public static byte[] DownloadImageBytes(this string url)
        {
            return _httpClient.GetByteArrayAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        public static byte[] ReadImageBytes(this string filePath)
        {
            return File.ReadAllBytes(filePath);
        }

        public static byte[] DecodeBase64ToBytes(this string base64String)
        {
            return Convert.FromBase64String(
                base64String.Replace("\r", "").Replace("\n", "").Replace(" ", ""));
        }

        private static string GetContentType(SKEncodedImageFormat format)
        {
            switch (format)
            {
                case SKEncodedImageFormat.Jpeg: return "image/jpeg";
                case SKEncodedImageFormat.Png: return "image/png";
                case SKEncodedImageFormat.Gif: return "image/gif";
                case SKEncodedImageFormat.Bmp: return "image/bmp";
                case SKEncodedImageFormat.Webp: return "image/webp";
                default: return "image/jpeg";
            }
        }

        #region 模板导出和图片导入仍在使用

        public static SixLabors.ImageSharp.Image GetImageByUrl(this string url, out SixLabors.ImageSharp.Formats.IImageFormat format)
        {
            using (Stream webStream = _httpClient.GetStreamAsync(url).ConfigureAwait(false).GetAwaiter().GetResult())
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    webStream.CopyTo(memoryStream);
                    memoryStream.Position = 0;

                    var image = SixLabors.ImageSharp.Image.Load(memoryStream);
                    format = image.GetImageFormat(memoryStream);

                    if (image.Metadata.HorizontalResolution <= 1 && image.Metadata.VerticalResolution <= 1)
                    {
                        image.Metadata.HorizontalResolution = SixLabors.ImageSharp.Metadata.ImageMetadata.DefaultHorizontalResolution;
                        image.Metadata.VerticalResolution = SixLabors.ImageSharp.Metadata.ImageMetadata.DefaultVerticalResolution;
                    }
                    return image;
                }
            }
        }

        public static SixLabors.ImageSharp.Image Base64StringToImage(this string base64String, out SixLabors.ImageSharp.Formats.IImageFormat format)
        {
            byte[] bytes = Convert.FromBase64String(
                base64String.Replace("\r", "").Replace("\n", "").Replace(" ", ""));
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                SixLabors.ImageSharp.Image image = SixLabors.ImageSharp.Image.Load(stream);
                format = image.GetImageFormat(stream);
                return image;
            }
        }

        public static string SaveTo(this SixLabors.ImageSharp.Image image, string path)
        {
            image.Save(path);
            return path;
        }

        public static string ToBase64String(this SixLabors.ImageSharp.Image image, SixLabors.ImageSharp.Formats.IImageFormat format)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, format);
                ms.Position = 0;
                return Convert.ToBase64String(ms.ToArray());
            }
        }

        #endregion
    }
}
