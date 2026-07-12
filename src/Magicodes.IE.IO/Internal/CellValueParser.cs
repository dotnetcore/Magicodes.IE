
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;

namespace Magicodes.IE.IO
{
    internal static class CellValueParser
    {
        private delegate object? ParseFunc(string cell);

        private static readonly Dictionary<Type, ParseFunc> _parsers = new()
        {
            [typeof(string)] = cell => cell,
            [typeof(int)] = cell => int.TryParse(cell, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null,
            [typeof(long)] = cell => long.TryParse(cell, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null,
            [typeof(double)] = cell => double.TryParse(cell, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v) ? v : null,
            [typeof(decimal)] = cell => decimal.TryParse(cell, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out var v) ? v : null,
            [typeof(bool)] = ParseBool,
            [typeof(DateTime)] = cell => ParseDateTime(cell),
            [typeof(DateTimeOffset)] = cell => ParseDateTimeOffset(cell),
            [typeof(Guid)] = cell => Guid.TryParse(cell, out var v) ? v : null,
            [typeof(ulong)] = ParseUInt64,
        };

        public static bool TryParse(string cell, Type targetType, out object? value, IReadOnlyList<CellConverter>? converters = null, bool date1904 = false)
        {
            if (converters is not null && converters.Count > 0)
            {
                for (int i = 0; i < converters.Count; i++)
                {
                    var conv = converters[i];
                    if (conv.Type == targetType)
                    {
                        var result = conv.TryRead(cell, out value);
                        if (!result) value = null;
                        return result;
                    }
                }
            }

            if (targetType == typeof(DateTime))
            {
                value = ParseDateTime(cell, date1904);
                return value is not null;
            }

            if (targetType == typeof(DateTimeOffset))
            {
                value = ParseDateTimeOffset(cell, date1904);
                return value is not null;
            }

            if (_parsers.TryGetValue(targetType, out var parser))
            {
                value = parser(cell);
                return value is not null;
            }

            if (targetType.IsEnum)
            {
                value = ParseEnum(cell, targetType);
                return value is not null;
            }

            try
            {
                value = Convert.ChangeType(cell, targetType, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex) when (ex is FormatException or OverflowException or InvalidCastException)
            {
                value = null;
                return false;
            }
        }

        private static object? ParseBool(string cell)
        {
            if (cell == "1" || cell.Equals("true", StringComparison.OrdinalIgnoreCase) || cell.Equals("yes", StringComparison.OrdinalIgnoreCase))
                return true;
            if (cell == "0" || cell.Equals("false", StringComparison.OrdinalIgnoreCase) || cell.Equals("no", StringComparison.OrdinalIgnoreCase))
                return false;
            return null;
        }

        private static object? ParseUInt64(string cell)
        {
            // Numbers are written with 'R' (round-trip) format, so large ulongs serialize as
            // scientific notation (e.g. "9.223372036854776E+18"). ulong.TryParse rejects the
            // exponent, so fall back to double then Convert — precision is bounded by double
            // anyway (Excel stores numbers as IEEE754 double).
            if (ulong.TryParse(cell, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)) return v;
            if (double.TryParse(cell, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d) && d >= 0)
            {
                try { return Convert.ToUInt64(d); } catch (OverflowException) { return null; }
            }
            return null;
        }

        private static object? ParseDateTime(string cell, bool date1904 = false)
        {
            if (double.TryParse(cell, NumberStyles.Any, CultureInfo.InvariantCulture, out var oa))
                return DateTime.FromOADate(oa + (date1904 ? 1462 : 0));
            if (DateTime.TryParse(cell, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dt))
                return dt;
            return null;
        }

        private static object? ParseDateTimeOffset(string cell, bool date1904 = false)
        {
            if (double.TryParse(cell, NumberStyles.Any, CultureInfo.InvariantCulture, out var oa))
                return new DateTimeOffset(DateTime.SpecifyKind(DateTime.FromOADate(oa + (date1904 ? 1462 : 0)), DateTimeKind.Unspecified));
            if (DateTimeOffset.TryParse(cell, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dto))
                return dto;
            return null;
        }

        private static object? ParseEnum(string cell, Type enumType)
        {
#if NETSTANDARD2_0
            try { return Enum.Parse(enumType, cell, ignoreCase: true); }
            catch { return null; }
#else
            return Enum.TryParse(enumType, cell, ignoreCase: true, out var ev) ? ev : null;
#endif
        }
    }
}