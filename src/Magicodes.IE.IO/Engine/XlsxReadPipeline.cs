
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Reflection;

namespace Magicodes.IE.IO
{
    internal static class XlsxReadPipeline
    {

        public static Func<int, PropertyInfo?> BuildResolver<T>(string[] headers, XlsxReadOptions<T>? profile) where T : new()
        {
            var type = typeof(T);
            var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var headerMap = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in props)
            {
                if (!p.CanWrite) continue;
                var importerAttr = p.GetCustomAttribute<ImporterHeaderAttribute>(inherit: true);
                var displayAttr = p.GetCustomAttribute<DisplayAttribute>(inherit: true);
                var descriptionAttr = p.GetCustomAttribute<DescriptionAttribute>(inherit: true);
                var name = importerAttr?.Name ?? displayAttr?.GetName() ?? descriptionAttr?.Description ?? p.Name;
                if (!headerMap.ContainsKey(name)) headerMap[name] = p;
                if (importerAttr?.Name is not null && !headerMap.ContainsKey(importerAttr.Name)) headerMap[importerAttr.Name] = p;
                if (displayAttr?.GetName() is { } displayName && !headerMap.ContainsKey(displayName)) headerMap[displayName] = p;
                if (descriptionAttr?.Description is { } description && !headerMap.ContainsKey(description)) headerMap[description] = p;
            }

            // An explicit column/header mapping wins over the names discovered from attributes.
            if (profile is not null)
            {
                return i =>
                {
                    var header = i < headers.Length ? headers[i] : null;
                    var configured = profile.Resolve(i, header);
                    if (configured is not null)
                        return headerMap.TryGetValue(configured, out var configuredProperty) ? configuredProperty : null;
                    if (headerMap.TryGetValue(header ?? "", out var p)) return p;
                    return null;
                };
            }

            return i => i < headers.Length && headerMap.TryGetValue(headers[i] ?? "", out var p) ? p : null;
        }

        public static Func<int, XlsxGeneratedPropertyMetadata<T>?> BuildGeneratedResolver<T>(
            string[] headers,
            XlsxReadOptions<T>? profile,
            IReadOnlyList<XlsxGeneratedPropertyMetadata<T>> metadata)
            where T : new()
        {
            var byName = new Dictionary<string, XlsxGeneratedPropertyMetadata<T>>(StringComparer.OrdinalIgnoreCase);
            var byHeader = new Dictionary<string, XlsxGeneratedPropertyMetadata<T>>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < metadata.Count; i++)
            {
                var property = metadata[i];
                if (property.Setter is null) continue;
                byName[property.Name] = property;
                if (!byHeader.ContainsKey(property.Name)) byHeader[property.Name] = property;
                if (property.ImportHeader is not null && !byHeader.ContainsKey(property.ImportHeader))
                    byHeader[property.ImportHeader] = property;
                if (property.DisplayName is not null && !byHeader.ContainsKey(property.DisplayName))
                    byHeader[property.DisplayName] = property;
                if (property.Description is not null && !byHeader.ContainsKey(property.Description))
                    byHeader[property.Description] = property;
            }

            return i =>
            {
                var header = i < headers.Length ? headers[i] : null;
                var configured = profile?.Resolve(i, header);
                if (configured is not null)
                    return byName.TryGetValue(configured, out var configuredProperty) ? configuredProperty : null;
                return byHeader.TryGetValue(header ?? string.Empty, out var property) ? property : null;
            };
        }

        public static void SetCellProperty(object item, PropertyInfo prop, string cell, Action<XlsxReadErrorInfo>? onParseError, int rowIndex, int colIndex, string? header, IReadOnlyList<CellConverter>? converters = null, bool date1904 = false)
        {
            var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

            if (CellValueParser.TryParse(cell, targetType, out var value, converters, date1904))
            {
                if (value is not null)
                    prop.SetValue(item, value);
            }
            else if (onParseError is not null)
            {
                onParseError(new XlsxReadErrorInfo
                {
                    RowIndex = rowIndex,
                    ColIndex = colIndex,
                    Header = header,
                    PropertyName = prop.Name,
                    RawCellValue = cell,
                    TargetTypeName = targetType.Name,
                    Exception = new FormatException($"Cannot convert '{cell}' to {targetType.Name}"),
                });
            }
        }

        public static void SetGeneratedProperty<T>(
            T item,
            XlsxGeneratedPropertyMetadata<T> property,
            string cell,
            Action<XlsxReadErrorInfo>? onParseError,
            int rowIndex,
            int colIndex,
            string? header,
            IReadOnlyList<CellConverter>? converters = null,
            bool date1904 = false)
            where T : new()
        {
            if (property.Setter is null || property.Setter(item, cell, converters, date1904))
                return;

            onParseError?.Invoke(new XlsxReadErrorInfo
            {
                RowIndex = rowIndex,
                ColIndex = colIndex,
                Header = header,
                PropertyName = property.Name,
                RawCellValue = cell,
                TargetTypeName = property.TargetTypeName,
                Exception = new FormatException($"Cannot convert '{cell}' to {property.TargetTypeName}"),
            });
        }
    }
}