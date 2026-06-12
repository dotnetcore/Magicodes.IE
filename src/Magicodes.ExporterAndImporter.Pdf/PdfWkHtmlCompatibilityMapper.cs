using System;
using WkHtmlToPdfDotNet;

namespace Magicodes.ExporterAndImporter.Pdf
{
    internal static class PdfWkHtmlCompatibilityMapper
    {
        internal static PdfExportOptions ToPdfExportOptions(PdfExporterAttribute attribute)
        {
            if (attribute == null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            return new PdfExportOptions
            {
                DocumentTitle = string.IsNullOrWhiteSpace(attribute.DocumentTitle) ? attribute.Name : attribute.DocumentTitle,
                Orientation = attribute.PageOrientation,
                PageSize = ToPdfPageSize(attribute),
                EnablePagesCount = attribute.IsEnablePagesCount,
                WriteHtml = attribute.IsWriteHtml,
                Header = Clone(attribute.Header),
                Footer = Clone(attribute.Footer),
                Margins = Clone(attribute.Margins)
            };
        }

        internal static PdfOrientation ToPdfOrientation(Orientation orientation)
        {
            return orientation == Orientation.Portrait ? PdfOrientation.Portrait : PdfOrientation.Landscape;
        }

        internal static Orientation ToWkHtmlOrientation(PdfOrientation orientation)
        {
            return orientation == PdfOrientation.Portrait ? Orientation.Portrait : Orientation.Landscape;
        }

        internal static PdfPageSize ToPdfPageSize(PaperKind paperKind, PdfCustomPaperSize customSize)
        {
            if (paperKind == PaperKind.Custom)
            {
                return new PdfPageSize
                {
                    IsCustom = true,
                    CustomSize = Clone(customSize)
                };
            }

            return PdfPageSize.Standard(paperKind.ToString());
        }

        internal static PdfPageSize ToPdfPageSize(PdfExporterAttribute attribute)
        {
            if (attribute == null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            if (attribute.CustomPageWidth > 0 && attribute.CustomPageHeight > 0)
            {
                return PdfPageSize.Custom(new PdfCustomPaperSize
                {
                    Width = attribute.CustomPageWidth,
                    Height = attribute.CustomPageHeight,
                    Unit = attribute.PageSizeUnit
                });
            }

            if (attribute.PageSize != null && attribute.PageSize.IsCustom && attribute.PageSize.CustomSize != null)
            {
                return Clone(attribute.PageSize);
            }

            if (!string.IsNullOrWhiteSpace(attribute.PageSizeName))
            {
                return PdfPageSize.Standard(attribute.PageSizeName);
            }

            if (attribute.PageSize != null)
            {
                return Clone(attribute.PageSize);
            }

            return PdfPageSize.Standard("A4");
        }

        internal static PaperKind ToWkHtmlPaperKind(PdfPageSize pageSize)
        {
            if (pageSize != null && pageSize.IsCustom)
            {
                return PaperKind.Custom;
            }

            if (pageSize != null && !string.IsNullOrWhiteSpace(pageSize.StandardName) &&
                Enum.TryParse(pageSize.StandardName, true, out PaperKind paperKind))
            {
                return paperKind;
            }

            return PaperKind.A4;
        }

        internal static PdfPageSize WithCustomPaperSize(PdfPageSize pageSize, PechkinPaperSize paperSize)
        {
            if (paperSize == null)
            {
                if (pageSize == null)
                {
                    return PdfPageSize.Standard("A4");
                }

                if (pageSize.IsCustom)
                {
                    return new PdfPageSize
                    {
                        IsCustom = true,
                        CustomSize = null
                    };
                }

                return Clone(pageSize);
            }

            return PdfPageSize.Custom(ToPdfCustomPaperSize(paperSize));
        }

        internal static PdfCustomPaperSize ToPdfCustomPaperSize(PechkinPaperSize paperSize)
        {
            if (paperSize == null)
            {
                return null;
            }

            return new PdfCustomPaperSize
            {
                Width = ParseLength(paperSize.Width).Value,
                Height = ParseLength(paperSize.Height).Value,
                Unit = ParseLength(paperSize.Width).Unit
            };
        }

        internal static PechkinPaperSize ToWkHtmlPaperSize(PdfPageSize pageSize)
        {
            if (pageSize == null || !pageSize.IsCustom || pageSize.CustomSize == null)
            {
                return null;
            }

            return new PechkinPaperSize(
                FormatLength(pageSize.CustomSize.Width, pageSize.CustomSize.Unit),
                FormatLength(pageSize.CustomSize.Height, pageSize.CustomSize.Unit));
        }

        internal static PdfHeaderOptions ToPdfHeaderOptions(HeaderSettings headerSettings)
        {
            if (headerSettings == null)
            {
                return null;
            }

            return new PdfHeaderOptions
            {
                FontSize = headerSettings.FontSize,
                FontName = headerSettings.FontName,
                Left = headerSettings.Left,
                Center = headerSettings.Center,
                Right = headerSettings.Right,
                Line = headerSettings.Line,
                Spacing = headerSettings.Spacing,
                HtmlUrl = headerSettings.HtmlUrl
            };
        }

        internal static HeaderSettings ToWkHtmlHeaderSettings(PdfHeaderOptions header)
        {
            if (header == null)
            {
                return null;
            }

            return new HeaderSettings
            {
                FontSize = header.FontSize,
                FontName = header.FontName,
                Left = header.Left,
                Center = header.Center,
                Right = header.Right,
                Line = header.Line,
                Spacing = header.Spacing,
                HtmlUrl = header.HtmlUrl
            };
        }

        internal static PdfFooterOptions ToPdfFooterOptions(FooterSettings footerSettings)
        {
            if (footerSettings == null)
            {
                return null;
            }

            return new PdfFooterOptions
            {
                FontSize = footerSettings.FontSize,
                FontName = footerSettings.FontName,
                Left = footerSettings.Left,
                Center = footerSettings.Center,
                Right = footerSettings.Right,
                Line = footerSettings.Line,
                Spacing = footerSettings.Spacing,
                HtmlUrl = footerSettings.HtmlUrl
            };
        }

        internal static FooterSettings ToWkHtmlFooterSettings(PdfFooterOptions footer)
        {
            if (footer == null)
            {
                return null;
            }

            return new FooterSettings
            {
                FontSize = footer.FontSize,
                FontName = footer.FontName,
                Left = footer.Left,
                Center = footer.Center,
                Right = footer.Right,
                Line = footer.Line,
                Spacing = footer.Spacing,
                HtmlUrl = footer.HtmlUrl
            };
        }

        internal static PdfMarginOptions ToPdfMarginOptions(MarginSettings marginSettings)
        {
            if (marginSettings == null)
            {
                return null;
            }

            return new PdfMarginOptions
            {
                Unit = ToPdfMeasurementUnit(marginSettings.Unit),
                Top = marginSettings.Top,
                Bottom = marginSettings.Bottom,
                Left = marginSettings.Left,
                Right = marginSettings.Right
            };
        }

        internal static MarginSettings ToWkHtmlMarginSettings(PdfMarginOptions margins)
        {
            if (margins == null)
            {
                return null;
            }

            return new MarginSettings
            {
                Unit = ToWkHtmlMeasurementUnit(margins.Unit),
                Top = margins.Top,
                Bottom = margins.Bottom,
                Left = margins.Left,
                Right = margins.Right
            };
        }

        internal static HtmlToPdfDocument ToHtmlToPdfDocument(PdfExportOptions options, string htmlString)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            var objSettings = new ObjectSettings
            {
                HtmlContent = htmlString,
                Encoding = System.Text.Encoding.UTF8,
                PagesCount = options.EnablePagesCount ? true : (bool?)null,
                WebSettings = { DefaultEncoding = System.Text.Encoding.UTF8.BodyName },
                HeaderSettings = ToWkHtmlHeaderSettings(options.Header),
                FooterSettings = ToWkHtmlFooterSettings(options.Footer)
            };

            var htmlToPdfDocument = new HtmlToPdfDocument
            {
                GlobalSettings =
                {
                    Orientation = ToWkHtmlOrientation(options.Orientation),
                    ColorMode = ColorMode.Color,
                    DocumentTitle = options.DocumentTitle
                },
                Objects =
                {
                    objSettings
                }
            };

            if (ToWkHtmlPaperKind(options.PageSize) == PaperKind.Custom)
            {
                htmlToPdfDocument.GlobalSettings.PaperSize = ToWkHtmlPaperSize(options.PageSize);
            }
            else
            {
                htmlToPdfDocument.GlobalSettings.PaperSize = ToWkHtmlPaperKind(options.PageSize);
            }

            var marginSettings = ToWkHtmlMarginSettings(options.Margins);
            if (marginSettings != null)
            {
                htmlToPdfDocument.GlobalSettings.Margins = marginSettings;
            }

            return htmlToPdfDocument;
        }

        private static PdfMeasurementUnit ToPdfMeasurementUnit(Unit unit)
        {
            switch (unit)
            {
                case Unit.Inches:
                    return PdfMeasurementUnit.Inches;
                case Unit.Centimeters:
                    return PdfMeasurementUnit.Centimeters;
                default:
                    return PdfMeasurementUnit.Millimeters;
            }
        }

        private static Unit ToWkHtmlMeasurementUnit(PdfMeasurementUnit unit)
        {
            switch (unit)
            {
                case PdfMeasurementUnit.Inches:
                    return Unit.Inches;
                case PdfMeasurementUnit.Centimeters:
                    return Unit.Centimeters;
                default:
                    return Unit.Millimeters;
            }
        }

        private static string FormatLength(double value, PdfMeasurementUnit unit)
        {
            var suffix = "mm";
            switch (unit)
            {
                case PdfMeasurementUnit.Inches:
                    suffix = "in";
                    break;
                case PdfMeasurementUnit.Centimeters:
                    suffix = "cm";
                    break;
            }

            return value.ToString(System.Globalization.CultureInfo.InvariantCulture) + suffix;
        }

        private static ParsedLength ParseLength(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return new ParsedLength(0d, PdfMeasurementUnit.Millimeters);
            }

            var trimmed = value.Trim();
            if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            {
                return new ParsedLength(ParseNumber(trimmed, 2), PdfMeasurementUnit.Centimeters);
            }

            if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            {
                return new ParsedLength(ParseNumber(trimmed, 2), PdfMeasurementUnit.Inches);
            }

            if (trimmed.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
            {
                return new ParsedLength(ParseNumber(trimmed, 2), PdfMeasurementUnit.Millimeters);
            }

            return new ParsedLength(ParseNumber(trimmed, 0), PdfMeasurementUnit.Millimeters);
        }

        private static double ParseNumber(string value, int unitLength)
        {
            var numericValue = unitLength > 0 ? value.Substring(0, value.Length - unitLength) : value;
            return double.Parse(numericValue, System.Globalization.CultureInfo.InvariantCulture);
        }

        private static PdfPageSize Clone(PdfPageSize pageSize)
        {
            if (pageSize == null)
            {
                return PdfPageSize.Standard("A4");
            }

            return new PdfPageSize
            {
                IsCustom = pageSize.IsCustom,
                StandardName = pageSize.StandardName,
                CustomSize = Clone(pageSize.CustomSize)
            };
        }

        private static PdfCustomPaperSize Clone(PdfCustomPaperSize customSize)
        {
            if (customSize == null)
            {
                return null;
            }

            return new PdfCustomPaperSize
            {
                Width = customSize.Width,
                Height = customSize.Height,
                Unit = customSize.Unit
            };
        }

        private static PdfHeaderOptions Clone(PdfHeaderOptions header)
        {
            if (header == null)
            {
                return null;
            }

            return new PdfHeaderOptions
            {
                FontSize = header.FontSize,
                FontName = header.FontName,
                Left = header.Left,
                Center = header.Center,
                Right = header.Right,
                Line = header.Line,
                Spacing = header.Spacing,
                HtmlUrl = header.HtmlUrl
            };
        }

        private static PdfFooterOptions Clone(PdfFooterOptions footer)
        {
            if (footer == null)
            {
                return null;
            }

            return new PdfFooterOptions
            {
                FontSize = footer.FontSize,
                FontName = footer.FontName,
                Left = footer.Left,
                Center = footer.Center,
                Right = footer.Right,
                Line = footer.Line,
                Spacing = footer.Spacing,
                HtmlUrl = footer.HtmlUrl
            };
        }

        private static PdfMarginOptions Clone(PdfMarginOptions margins)
        {
            if (margins == null)
            {
                return null;
            }

            return new PdfMarginOptions
            {
                Unit = margins.Unit,
                Top = margins.Top,
                Bottom = margins.Bottom,
                Left = margins.Left,
                Right = margins.Right
            };
        }

        private readonly struct ParsedLength
        {
            internal ParsedLength(double value, PdfMeasurementUnit unit)
            {
                Value = value;
                Unit = unit;
            }

            internal double Value { get; }

            internal PdfMeasurementUnit Unit { get; }
        }
    }
}
