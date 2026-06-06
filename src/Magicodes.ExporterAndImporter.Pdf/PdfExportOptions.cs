using System;

namespace Magicodes.ExporterAndImporter.Pdf
{
    public class PdfExportOptions
    {
        public string DocumentTitle { get; set; }

        public PdfOrientation Orientation { get; set; } = PdfOrientation.Landscape;

        public PdfPageSize PageSize { get; set; } = PdfPageSize.Standard("A4");

        public bool EnablePagesCount { get; set; }

        public bool WriteHtml { get; set; }

        public PdfHeaderOptions Header { get; set; }

        public PdfFooterOptions Footer { get; set; }

        public PdfMarginOptions Margins { get; set; }
    }

    public enum PdfOrientation
    {
        Landscape = 0,
        Portrait = 1
    }

    public class PdfPageSize
    {
        public bool IsCustom { get; set; }

        public string StandardName { get; set; } = "A4";

        public PdfCustomPaperSize CustomSize { get; set; }

        public static PdfPageSize Standard(string standardName)
        {
            if (string.IsNullOrWhiteSpace(standardName))
            {
                throw new ArgumentException("A standard page size name is required.", nameof(standardName));
            }

            return new PdfPageSize
            {
                IsCustom = false,
                StandardName = standardName,
                CustomSize = null
            };
        }

        public static PdfPageSize Custom(PdfCustomPaperSize customSize)
        {
            if (customSize == null)
            {
                throw new ArgumentNullException(nameof(customSize));
            }

            return new PdfPageSize
            {
                IsCustom = true,
                StandardName = null,
                CustomSize = customSize
            };
        }
    }

    public class PdfCustomPaperSize
    {
        public double Width { get; set; }

        public double Height { get; set; }

        public PdfMeasurementUnit Unit { get; set; } = PdfMeasurementUnit.Millimeters;
    }

    public enum PdfMeasurementUnit
    {
        Inches = 0,
        Millimeters = 1,
        Centimeters = 2
    }

    public class PdfHeaderOptions
    {
        public int? FontSize { get; set; }

        public string FontName { get; set; }

        public string Left { get; set; }

        public string Center { get; set; }

        public string Right { get; set; }

        public bool? Line { get; set; }

        public double? Spacing { get; set; }

        public string HtmlUrl { get; set; }
    }

    public class PdfFooterOptions
    {
        public int? FontSize { get; set; }

        public string FontName { get; set; }

        public string Left { get; set; }

        public string Center { get; set; }

        public string Right { get; set; }

        public bool? Line { get; set; }

        public double? Spacing { get; set; }

        public string HtmlUrl { get; set; }
    }

    public class PdfMarginOptions
    {
        public PdfMeasurementUnit Unit { get; set; } = PdfMeasurementUnit.Millimeters;

        public double? Top { get; set; }

        public double? Bottom { get; set; }

        public double? Left { get; set; }

        public double? Right { get; set; }
    }
}
