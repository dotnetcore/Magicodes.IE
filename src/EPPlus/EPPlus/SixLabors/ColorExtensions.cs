using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Magicodes.IE.EPPlus.SixLabors
{
    internal static class ColorExtensions
    {
        public static string ToArgbHex(this Color color)
        {
            var rgba = color.ToPixel<Rgba32>();
            return rgba.A.ToString("X2") + rgba.R.ToString("X2") + rgba.G.ToString("X2") + rgba.B.ToString("X2");
        }
    }
}