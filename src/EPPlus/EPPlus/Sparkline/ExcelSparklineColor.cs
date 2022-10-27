using OfficeOpenXml.Style;
using System;
using System.Globalization;
using System.Xml;
using Magicodes.IE.EPPlus.SixLabors;
using SixLabors.ImageSharp;

namespace OfficeOpenXml.Sparkline
{
    /// <summary>
    /// Sparkline colors
    /// </summary>
    public class ExcelSparklineColor : XmlHelper, IColor
    {
        internal ExcelSparklineColor(XmlNamespaceManager ns, XmlNode node) : base(ns, node)
        {

        }
        /// <summary>
        /// Indexed color
        /// </summary>
        public int Indexed
        {
            get => GetXmlNodeInt("@indexed");
            set
            {
                if (value < 0 || value > 65)
                {
                    throw (new ArgumentOutOfRangeException("Index out of range"));
                }

                SetXmlNodeString("@indexed", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// RGB 
        /// </summary>
        public string Rgb
        {
            get => GetXmlNodeString("@rgb");
            internal set
            {
                SetXmlNodeString("@rgb", value);
            }
        }


        public string Theme => GetXmlNodeString("@theme");

        /// <summary>
        /// The tint value
        /// </summary>
        public decimal Tint
        {
            get => GetXmlNodeDecimal("@tint");
            set
            {
                if (value > 1 || value < -1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between -1 and 1"));
                }
                SetXmlNodeString("@tint", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Sets a color
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor(Color color)
        {
            Rgb = color.ToArgbHex();
        }
    }
}
