
using System;
using System.Buffers;
using System.Globalization;
using System.Text;

namespace Magicodes.IE.IO
{
    internal static class NumberFormatHelper
    {

        public static void WriteInt32(int value, Span<byte> dest, out int written)
        {
#if NETSTANDARD2_0
            var s = value.ToString(CultureInfo.InvariantCulture);
            var tmp = Encoding.UTF8.GetBytes(s);
            if (tmp.Length > dest.Length)
                throw new InvalidOperationException($"NumberFormatHelper.WriteInt32: buffer too small (need {tmp.Length}, have {dest.Length})");
            tmp.AsSpan().CopyTo(dest);
            written = tmp.Length;
#else
            if (!System.Buffers.Text.Utf8Formatter.TryFormat(value, dest, out written))
                throw new InvalidOperationException("Utf8Formatter.TryFormat failed (buffer too small?)");
#endif
        }

        public static void WriteInt64(long value, Span<byte> dest, out int written)
        {
#if NETSTANDARD2_0
            var s = value.ToString(CultureInfo.InvariantCulture);
            var tmp = Encoding.UTF8.GetBytes(s);
            if (tmp.Length > dest.Length)
                throw new InvalidOperationException($"NumberFormatHelper.WriteInt64: buffer too small (need {tmp.Length}, have {dest.Length})");
            tmp.AsSpan().CopyTo(dest);
            written = tmp.Length;
#else
            if (!System.Buffers.Text.Utf8Formatter.TryFormat(value, dest, out written))
                throw new InvalidOperationException("Utf8Formatter.TryFormat failed (buffer too small?)");
#endif
        }

        public static void WriteDouble(double value, char format, Span<byte> dest, out int written)
        {
            if (format != 'R' && format != 'r' && format != 'G' && format != 'g' && format != 'F' && format != 'f')
                throw new ArgumentException($"WriteDouble: unsupported format '{format}' (only R/G/F accepted)", nameof(format));
#if NETSTANDARD2_0
            // On .NET Framework, double.ToString("R") emits 15 sig digits and is NOT reliably
            // round-trippable (the shortest-round-trip "R" fix landed in .NET Core 3.0).
            // "G17" always emits 17 sig digits → uniquely identifies the double on every TFM.
            // The modern path uses Utf8Formatter 'R' (shortest round-trip); both round-trip
            // correctly, the ns2.0 output is just more verbose (17 digits vs shortest).
            var fmt = (format == 'R' || format == 'r') ? "G17" : (format == 'F' || format == 'f') ? "F" : "G";
            var s = value.ToString(fmt, CultureInfo.InvariantCulture);
            var tmp = Encoding.UTF8.GetBytes(s);
            if (tmp.Length > dest.Length)
                throw new InvalidOperationException($"NumberFormatHelper.WriteDouble: buffer too small (need {tmp.Length}, have {dest.Length})");
            tmp.AsSpan().CopyTo(dest);
            written = tmp.Length;
#else
            if (!System.Buffers.Text.Utf8Formatter.TryFormat(value, dest, out written, new System.Buffers.StandardFormat(format)))
                throw new InvalidOperationException("Utf8Formatter.TryFormat failed (buffer too small?)");
#endif
        }

        public static void WriteDoubleFixedTrimmed(double value, int decimals, Span<byte> dest, out int written)
        {
            if (decimals < 0 || decimals > 28)
                throw new ArgumentOutOfRangeException(nameof(decimals));

            value = Math.Round(value, decimals, MidpointRounding.AwayFromZero);

#if NETSTANDARD2_0
            var s = value.ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            var tmp = Encoding.UTF8.GetBytes(s);
            if (tmp.Length > dest.Length)
                throw new InvalidOperationException($"NumberFormatHelper.WriteDoubleFixedTrimmed: buffer too small (need {tmp.Length}, have {dest.Length})");
            tmp.AsSpan().CopyTo(dest);
            written = tmp.Length;
#else
            if (!System.Buffers.Text.Utf8Formatter.TryFormat(value, dest, out written, new System.Buffers.StandardFormat('F', (byte)decimals)))
                throw new InvalidOperationException("Utf8Formatter.TryFormat failed (buffer too small?)");
#endif

            while (written > 0 && dest[written - 1] == (byte)'0')
                written--;
            if (written > 0 && dest[written - 1] == (byte)'.')
                written--;
        }
    }
}
