
using System;
using System.Buffers;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Magicodes.IE.IO
{
    internal sealed class ByteBufferWriter : IBufferWriter<byte>, IDisposable
    {
        private byte[] _buffer;
        private int _pos;
        private const int MinChunkSize = 8 * 1024;

        public ByteBufferWriter(int initialSize)
        {
            if (initialSize < 256) initialSize = 256;
            _buffer = ArrayPool<byte>.Shared.Rent(initialSize);
            _pos = 0;
        }

        public int WrittenCount => _pos;

        public void Advance(int count) => _pos += count;

        public Memory<byte> GetMemory(int sizeHint = 0)
        {
            EnsureCapacity(sizeHint);
            return _buffer.AsMemory(_pos);
        }

        public Span<byte> GetSpan(int sizeHint = 0)
        {
            EnsureCapacity(sizeHint);
            return _buffer.AsSpan(_pos);
        }

        private void EnsureCapacity(int sizeHint)
        {
            int need = sizeHint == 0 ? MinChunkSize : sizeHint;
            if (_pos + need <= _buffer.Length) return;
            int newSize = Math.Max(_buffer.Length * 2, _pos + need);
            var nb = ArrayPool<byte>.Shared.Rent(newSize);
            Buffer.BlockCopy(_buffer, 0, nb, 0, _pos);
            ArrayPool<byte>.Shared.Return(_buffer);
            _buffer = nb;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void WriteUtf8(ReadOnlySpan<byte> utf8)
        {
            int len = utf8.Length;
            if (_pos + len > _buffer.Length)
            {
                EnsureCapacity(len);
            }
            utf8.CopyTo(_buffer.AsSpan(_pos));
            _pos += len;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void WriteUtf8(string s)
        {
            if (string.IsNullOrEmpty(s)) return;
#if NETSTANDARD2_0
            WriteUtf8(System.Text.Encoding.UTF8.GetBytes(s));
#else
            var span = s.AsSpan();
            int need = System.Text.Encoding.UTF8.GetByteCount(span);
            if (_pos + need > _buffer.Length) EnsureCapacity(need);
            System.Text.Encoding.UTF8.GetBytes(span, _buffer.AsSpan(_pos));
            _pos += need;
#endif
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void WriteUtf8(System.Text.StringBuilder sb)
        {
            int len = sb.Length;
            if (len == 0) return;
#if NETSTANDARD2_0
            int maxBytes = System.Text.Encoding.UTF8.GetMaxByteCount(len);
            byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(maxBytes);
            try
            {
                int n = System.Text.Encoding.UTF8.GetBytes(sb.ToString(), 0, len, buf, 0);
                if (_pos + n > _buffer.Length) EnsureCapacity(n);
                buf.AsSpan(0, n).CopyTo(_buffer.AsSpan(_pos));
                _pos += n;
            }
            finally
            {
                System.Buffers.ArrayPool<byte>.Shared.Return(buf);
            }
#else
            int totalBytes = 0;
            foreach (var chunk in sb.GetChunks())
            {
                totalBytes += System.Text.Encoding.UTF8.GetByteCount(chunk.Span);
            }
            if (_pos + totalBytes > _buffer.Length) EnsureCapacity(totalBytes);
            var dst = _buffer.AsSpan(_pos, totalBytes);
            int written = 0;
            foreach (var chunk in sb.GetChunks())
            {
                written += System.Text.Encoding.UTF8.GetBytes(chunk.Span, dst.Slice(written));
            }
            _pos += written;
#endif
        }

        public void WriteEscaped(string s)
        {
            int spanLen = s.Length;
            if (spanLen == 0) return;

            int i = 0;
            int plainStart = 0;
            while (i < spanLen)
            {
                char c = s[i];
                if (IsIllegalXmlChar(c))
                {
                    FlushPlain(s, plainStart, i);
                    WriteUtf8(ReplacementUtf8);
                    i++;
                    plainStart = i;
                    continue;
                }
                if (c >= 0x20 || c == '\t' || c == '\n' || c == '\r')
                {
                    if (c < 0x80)
                    {
                        int escLen = GetAsciiEscapeLen(c);
                        if (escLen > 0)
                        {
                            FlushPlain(s, plainStart, i);
                            WriteAsciiEscape(c);
                            i++;
                            plainStart = i;
                            continue;
                        }
                        i++;
                    }
                    else
                    {
                        i++;
                    }
                }
                else
                {
                    FlushPlain(s, plainStart, i);
                    if (_pos >= _buffer.Length) EnsureCapacity(1);
                    _buffer[_pos++] = 0x20;
                    i++;
                    plainStart = i;
                }
            }
            FlushPlain(s, plainStart, spanLen);
        }

        private void FlushPlain(string s, int start, int end)
        {
            if (end <= start) return;
#if NET8_0_OR_GREATER
            if (end - start >= VectorEscapeThreshold)
            {
                var vspan = s.AsSpan(start, end - start);
                if (System.Text.Ascii.IsValid(vspan))
                {
                    int vlen = vspan.Length;
                    if (_pos + vlen > _buffer.Length) EnsureCapacity(vlen);
                    System.Text.Ascii.FromUtf16(vspan, _buffer.AsSpan(_pos, vlen), out int vnb);
                    _pos += vnb;
                    return;
                }
                int vneed = Encoding.UTF8.GetByteCount(vspan);
                if (_pos + vneed > _buffer.Length) EnsureCapacity(vneed);
                _pos += Encoding.UTF8.GetBytes(vspan, _buffer.AsSpan(_pos));
                return;
            }
#endif
            bool asciiOnly = true;
            for (int j = start; j < end; j++)
            {
                if (s[j] >= 0x80) { asciiOnly = false; break; }
            }
            if (asciiOnly)
            {
                int len = end - start;
                if (_pos + len > _buffer.Length) EnsureCapacity(len);
                for (int j = start; j < end; j++) _buffer[_pos++] = (byte)s[j];
            }
            else
            {
                var span = s.AsSpan(start, end - start);
#if NETSTANDARD2_0
                int need = Encoding.UTF8.GetByteCount(span.ToString());
#else
                int need = Encoding.UTF8.GetByteCount(span);
#endif
                if (_pos + need > _buffer.Length) EnsureCapacity(need);
#if NETSTANDARD2_0
                var tmp = Encoding.UTF8.GetBytes(s.Substring(start, end - start));
                tmp.AsSpan().CopyTo(_buffer.AsSpan(_pos));
                _pos += tmp.Length;
#else
                int n = Encoding.UTF8.GetBytes(span, _buffer.AsSpan(_pos));
                _pos += n;
#endif
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int GetAsciiEscapeLen(char c)
        {
            if (c == '&') return 5;
            if (c == '<') return 4;
            if (c == '>') return 4;
            if (c == '"') return 6;
            if (c == '\'') return 6;
            return 0;
        }

        private static readonly byte[] ReplacementUtf8 = Encoding.UTF8.GetBytes("\uFFFD");

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static bool IsIllegalXmlChar(char c) =>
            (c < 0x20 && c != '\t' && c != '\n' && c != '\r') || c is '\uFFFE' or '\uFFFF';

        private static ReadOnlySpan<byte> AmpEsc => new byte[] { (byte)'&', (byte)'a', (byte)'m', (byte)'p', (byte)';' };
        private static ReadOnlySpan<byte> LtEsc  => new byte[] { (byte)'&', (byte)'l', (byte)'t', (byte)';' };
        private static ReadOnlySpan<byte> GtEsc  => new byte[] { (byte)'&', (byte)'g', (byte)'t', (byte)';' };
        private static ReadOnlySpan<byte> QuotEsc => new byte[] { (byte)'&', (byte)'q', (byte)'u', (byte)'o', (byte)'t', (byte)';' };
        private static ReadOnlySpan<byte> AposEsc => new byte[] { (byte)'&', (byte)'a', (byte)'p', (byte)'o', (byte)'s', (byte)';' };

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void WriteAsciiEscape(char c)
        {
            ReadOnlySpan<byte> esc = c switch
            {
                '&' => AmpEsc,
                '<' => LtEsc,
                '>' => GtEsc,
                '"' => QuotEsc,
                '\'' => AposEsc,
                _ => default,
            };
            int len = esc.Length;
            if (_pos + len > _buffer.Length) EnsureCapacity(len);
            esc.CopyTo(_buffer.AsSpan(_pos));
            _pos += len;
        }

        public void WriteDouble(double d)
        {
            Span<byte> buf = stackalloc byte[32];
            NumberFormatHelper.WriteDouble(d, 'R', buf, out int w);
            if (_pos + w > _buffer.Length) EnsureCapacity(w);
            buf.Slice(0, w).CopyTo(_buffer.AsSpan(_pos));
            _pos += w;
        }

        public void WriteNumberCell(int styleId, ReadOnlySpan<byte> colLetter, ReadOnlySpan<byte> rowRef, double d, bool strictCellReferences = true)
        {
            if (!strictCellReferences)
            {
                WriteUtf8("<c t=\"n\""u8);
                if (styleId > 0)
                {
                    WriteUtf8(" s=\""u8);
                    WriteInt32(styleId);
                    WriteUtf8("\""u8);
                }
                WriteUtf8("><v>"u8);
                WriteDouble(d);
                WriteUtf8("</v></c>"u8);
                return;
            }
            int maxBytes = 73 + colLetter.Length;
            if (_pos + maxBytes > _buffer.Length) EnsureCapacity(maxBytes);

            var dst = _buffer.AsSpan(_pos, maxBytes);
            int p = 0;
            dst[p++] = (byte)'<';
            dst[p++] = (byte)'c';
            dst[p++] = (byte)' ';
            dst[p++] = (byte)'r';
            dst[p++] = (byte)'=';
            dst[p++] = (byte)'"';
            colLetter.CopyTo(dst.Slice(p));
            p += colLetter.Length;
            rowRef.CopyTo(dst.Slice(p));
            p += rowRef.Length;
            dst[p++] = (byte)'"';
            dst[p++] = (byte)' ';
            dst[p++] = (byte)'t';
            dst[p++] = (byte)'=';
            dst[p++] = (byte)'"';
            dst[p++] = (byte)'n';
            dst[p++] = (byte)'"';
            if (styleId > 0)
            {
                dst[p++] = (byte)' ';
                dst[p++] = (byte)'s';
                dst[p++] = (byte)'=';
                dst[p++] = (byte)'"';
                NumberFormatHelper.WriteInt32(styleId, dst.Slice(p), out int sw);
                p += sw;
                dst[p++] = (byte)'"';
            }
            dst[p++] = (byte)'>';
            dst[p++] = (byte)'<';
            dst[p++] = (byte)'v';
            dst[p++] = (byte)'>';
            NumberFormatHelper.WriteDouble(d, 'R', dst.Slice(p), out int dw);
            p += dw;
            dst[p++] = (byte)'<';
            dst[p++] = (byte)'/';
            dst[p++] = (byte)'v';
            dst[p++] = (byte)'>';
            dst[p++] = (byte)'<';
            dst[p++] = (byte)'/';
            dst[p++] = (byte)'c';
            dst[p++] = (byte)'>';

            _pos += p;
        }

        public void WriteInt32(int n)
        {
            Span<byte> buf = stackalloc byte[12];
            NumberFormatHelper.WriteInt32(n, buf, out int w);
            if (_pos + w > _buffer.Length) EnsureCapacity(w);
            buf.Slice(0, w).CopyTo(_buffer.AsSpan(_pos));
            _pos += w;
        }

        public void WriteInt64(long n)
        {
            Span<byte> buf = stackalloc byte[20];
            NumberFormatHelper.WriteInt64(n, buf, out int w);
            if (_pos + w > _buffer.Length) EnsureCapacity(w);
            buf.Slice(0, w).CopyTo(_buffer.AsSpan(_pos));
            _pos += w;
        }

        public void WriteInlineStringCell(int styleId, string? s, ReadOnlySpan<byte> colLetter, ReadOnlySpan<byte> rowRef, bool strictCellReferences = true)
        {
            if (!strictCellReferences)
            {
                WriteInlineStringCellFallback(styleId, s ?? string.Empty, ReadOnlySpan<byte>.Empty, ReadOnlySpan<byte>.Empty);
                return;
            }
            if (s is null)
            {
                WriteInlineStringCellFallback(styleId, "", colLetter, rowRef);
                return;
            }

            int sLen = s.Length;

            if (sLen > 0 && IsAsciiNoEscape(s.AsSpan())
                && !(char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[sLen - 1])))
            {
                int exactBytes = InlineStringCellPrefixLen
                    + (strictCellReferences ? colLetter.Length + rowRef.Length : 0)
                    + sLen
                    + InlineStringCellSuffixLen;
                if (styleId > 0)
                {
                    exactBytes += StyleAttrOpeningLen + CountDigits(styleId) + StyleAttrClosingQuoteLen;
                }
                if (_pos + exactBytes <= _buffer.Length)
                {
                    var dst = _buffer.AsSpan(_pos, exactBytes);
                    int p = 0;
                    ReadOnlySpan<byte> prefixHead = "<c"u8;
                    prefixHead.CopyTo(dst.Slice(p)); p += prefixHead.Length;
                    if (strictCellReferences)
                    {
                        ReadOnlySpan<byte> rOpen = " r=\""u8;
                        rOpen.CopyTo(dst.Slice(p)); p += rOpen.Length;
                        colLetter.CopyTo(dst.Slice(p)); p += colLetter.Length;
                        rowRef.CopyTo(dst.Slice(p)); p += rowRef.Length;
                        dst[p++] = (byte)'"';
                    }
                    ReadOnlySpan<byte> prefixTail = " t=\"inlineStr\""u8;
                    prefixTail.CopyTo(dst.Slice(p)); p += prefixTail.Length;
                    if (styleId > 0)
                    {
                        ReadOnlySpan<byte> sOpen = " s=\""u8;
                        sOpen.CopyTo(dst.Slice(p)); p += sOpen.Length;
                        NumberFormatHelper.WriteInt32(styleId, dst.Slice(p), out int sw); p += sw;
                        dst[p++] = (byte)'"';
                    }
                    ReadOnlySpan<byte> afterStyle = "><is><t>"u8;
                    afterStyle.CopyTo(dst.Slice(p)); p += afterStyle.Length;
#if NET8_0_OR_GREATER
                    if (sLen >= VectorEscapeThreshold)
                    {
                        System.Text.Ascii.FromUtf16(s.AsSpan(), dst.Slice(p, sLen), out int narrowed);
                        p += narrowed;
                    }
                    else
                    {
                        for (int i = 0; i < sLen; i++) dst[p++] = (byte)s[i];
                    }
#else
                    for (int i = 0; i < sLen; i++) dst[p++] = (byte)s[i];
#endif
                    ReadOnlySpan<byte> suffix = "</t></is></c>"u8;
                    suffix.CopyTo(dst.Slice(p)); p += suffix.Length;
                    _pos += p;
                    return;
                }
                WriteInlineStringCellFallback(styleId, s, colLetter, rowRef);
                return;
            }

            long maxBytes = 96L + colLetter.Length + rowRef.Length + (long)sLen * 6;
            if (maxBytes > int.MaxValue || maxBytes > _buffer.Length - _pos)
            {
                WriteInlineStringCellFallback(styleId, s, colLetter, rowRef);
                return;
            }

            var dst2 = _buffer.AsSpan(_pos, (int)maxBytes);
            int p2 = 0;

            dst2[p2++] = (byte)'<';
            dst2[p2++] = (byte)'c';
            if (!colLetter.IsEmpty)
            {
                dst2[p2++] = (byte)' ';
                dst2[p2++] = (byte)'r';
                dst2[p2++] = (byte)'=';
                dst2[p2++] = (byte)'"';
                colLetter.CopyTo(dst2.Slice(p2));
                p2 += colLetter.Length;
                rowRef.CopyTo(dst2.Slice(p2)); p2 += rowRef.Length;
                dst2[p2++] = (byte)'"';
            }
            dst2[p2++] = (byte)' ';
            dst2[p2++] = (byte)'t';
            dst2[p2++] = (byte)'=';
            dst2[p2++] = (byte)'"';
            dst2[p2++] = (byte)'i';
            dst2[p2++] = (byte)'n';
            dst2[p2++] = (byte)'l';
            dst2[p2++] = (byte)'i';
            dst2[p2++] = (byte)'n';
            dst2[p2++] = (byte)'e';
            dst2[p2++] = (byte)'S';
            dst2[p2++] = (byte)'t';
            dst2[p2++] = (byte)'r';
            dst2[p2++] = (byte)'"';
            if (styleId > 0)
            {
                dst2[p2++] = (byte)' ';
                dst2[p2++] = (byte)'s';
                dst2[p2++] = (byte)'=';
                dst2[p2++] = (byte)'"';
                NumberFormatHelper.WriteInt32(styleId, dst2.Slice(p2), out int sw2); p2 += sw2;
                dst2[p2++] = (byte)'"';
            }
            dst2[p2++] = (byte)'>';

            dst2[p2++] = (byte)'<';
            dst2[p2++] = (byte)'i';
            dst2[p2++] = (byte)'s';
            dst2[p2++] = (byte)'>';
            dst2[p2++] = (byte)'<';
            dst2[p2++] = (byte)'t';
            bool preserve = sLen > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[sLen - 1]));
            if (preserve)
            {
                ReadOnlySpan<byte> xs = " xml:space=\"preserve\""u8;
                xs.CopyTo(dst2.Slice(p2));
                p2 += xs.Length;
            }
            dst2[p2++] = (byte)'>';

            p2 = WriteEscapedInto(s, 0, sLen, dst2, p2);

            ReadOnlySpan<byte> tail = "</t></is></c>"u8;
            tail.CopyTo(dst2.Slice(p2));
            p2 += tail.Length;

            _pos += p2;
        }

        private static readonly int InlineStringCellPrefixLen = "<c r=\"".Length + "\" t=\"inlineStr\"".Length + "><is><t>".Length;
        private static readonly int InlineStringCellSuffixLen = "</t></is></c>".Length;
        private static readonly int StyleAttrOpeningLen = " s=\"".Length;
        private static readonly int StyleAttrClosingQuoteLen = "\"".Length;

#if NET8_0_OR_GREATER
        private static readonly System.Buffers.SearchValues<char> SafeInlineChars =
            System.Buffers.SearchValues.Create(BuildSafeInlineChars());

        private const int VectorEscapeThreshold = 32;

        private static string BuildSafeInlineChars()
        {
            var sb = new StringBuilder(128);
            sb.Append('\t').Append('\n').Append('\r');
            for (int c = 0x20; c < 0x80; c++)
            {
                if (c == '&' || c == '<' || c == '>' || c == '"' || c == '\'') continue;
                sb.Append((char)c);
            }
            return sb.ToString();
        }
#endif

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static bool IsAsciiNoEscape(ReadOnlySpan<char> s)
        {
#if NET8_0_OR_GREATER
            if (s.Length >= VectorEscapeThreshold)
                return s.IndexOfAnyExcept(SafeInlineChars) < 0;
#endif
            int len = s.Length;
            for (int i = 0; i < len; i++)
            {
                char c = s[i];
                if (c >= 0x80) return false;
                if (c < 0x20 && c != '\t' && c != '\n' && c != '\r') return false;
                if (c == '&' || c == '<' || c == '>' || c == '"' || c == '\'') return false;
            }
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int CountDigits(int n)
        {
            if (n < 10) return 1;
            if (n < 100) return 2;
            if (n < 1000) return 3;
            if (n < 10000) return 4;
            if (n < 100000) return 5;
            return 6;
        }

        private void WriteInlineStringCellFallback(int styleId, string s, ReadOnlySpan<byte> colLetter, ReadOnlySpan<byte> rowRef)
        {
            WriteUtf8("<c"u8);
            if (!colLetter.IsEmpty)
            {
                WriteUtf8(" r=\""u8);
                WriteUtf8(colLetter);
                WriteUtf8(rowRef);
                WriteUtf8("\""u8);
            }
            WriteUtf8(" t=\"inlineStr\""u8);
            if (styleId > 0)
            {
                WriteUtf8(" s=\""u8);
                WriteInt32(styleId);
                WriteUtf8("\""u8);
            }
            WriteUtf8("><is><t"u8);
            int sLen = s.Length;
            bool preserve = sLen > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[sLen - 1]));
            if (preserve) WriteUtf8(" xml:space=\"preserve\""u8);
            WriteUtf8(">"u8);
            WriteEscaped(s);
            WriteUtf8("</t></is></c>"u8);
        }

        private static int WriteEscapedInto(string s, int start, int end, Span<byte> dst, int p)
        {
            int i = start;
            int plainStart = start;
            int cap = dst.Length;
            while (i < end)
            {
                char c = s[i];
                if (IsIllegalXmlChar(c))
                {
                    p = FlushPlainInto(s, plainStart, i, dst, p, cap);
                    if (p + ReplacementUtf8.Length > cap) return p;
                    ReplacementUtf8.AsSpan().CopyTo(dst.Slice(p));
                    p += ReplacementUtf8.Length;
                    i++;
                    plainStart = i;
                    continue;
                }
                if (c >= 0x20 || c == '\t' || c == '\n' || c == '\r')
                {
                    if (c < 0x80)
                    {
                        int escLen = GetAsciiEscapeLen(c);
                        if (escLen > 0)
                        {
                            p = FlushPlainInto(s, plainStart, i, dst, p, cap);
                            ReadOnlySpan<byte> esc = c switch
                            {
                                '&' => AmpEsc,
                                '<' => LtEsc,
                                '>' => GtEsc,
                                '"' => QuotEsc,
                                '\'' => AposEsc,
                                _ => default,
                            };
                            esc.CopyTo(dst.Slice(p));
                            p += esc.Length;
                            i++;
                            plainStart = i;
                            continue;
                        }
                        i++;
                    }
                    else
                    {
                        i++;
                    }
                }
                else
                {
                    p = FlushPlainInto(s, plainStart, i, dst, p, cap);
                    if (p >= cap) return p;
                    dst[p++] = (byte)0x20;
                    i++;
                    plainStart = i;
                }
            }
            return FlushPlainInto(s, plainStart, end, dst, p, cap);
        }

        private static int FlushPlainInto(string s, int start, int end, Span<byte> dst, int p, int cap)
        {
            if (end <= start) return p;
#if NET8_0_OR_GREATER
            if (end - start >= VectorEscapeThreshold)
            {
                var vspan = s.AsSpan(start, end - start);
                if (System.Text.Ascii.IsValid(vspan))
                {
                    int vlen = vspan.Length;
                    if (p + vlen > cap) return p;
                    System.Text.Ascii.FromUtf16(vspan, dst.Slice(p, vlen), out int vnb);
                    return p + vnb;
                }
                if (p >= cap) return p;
                return p + Encoding.UTF8.GetBytes(vspan, dst.Slice(p));
            }
#endif
            bool asciiOnly = true;
            for (int j = start; j < end; j++)
            {
                if (s[j] >= 0x80) { asciiOnly = false; break; }
            }
            if (asciiOnly)
            {
                int len = end - start;
                if (p + len > cap) return p;
                for (int j = start; j < end; j++) dst[p++] = (byte)s[j];
                return p;
            }
            else
            {
                var span = s.AsSpan(start, end - start);
                if (p >= cap) return p;
                var destAvail = dst.Slice(p);
#if NETSTANDARD2_0
                int n2 = Encoding.UTF8.GetByteCount(span.ToString());
                if (n2 > destAvail.Length) return p;
                var tmp = Encoding.UTF8.GetBytes(span.ToString());
                tmp.AsSpan().CopyTo(destAvail);
                return p + tmp.Length;
#else
                return p + Encoding.UTF8.GetBytes(span, destAvail);
#endif
            }
        }

        public int RemainingCapacity => _buffer.Length - _pos;

        public void FlushTo(Stream stream)
        {
            if (_pos == 0) return;
            stream.Write(_buffer, 0, _pos);
            _pos = 0;
        }

        public ValueTask FlushToAsync(Stream stream, CancellationToken cancellationToken = default)
        {
            if (_pos == 0) return default;
#if NETSTANDARD2_0
            return FlushToAsyncCoreNs20(stream, cancellationToken);
#else
            var vt = stream.WriteAsync((ReadOnlyMemory<byte>)_buffer.AsMemory(0, _pos), cancellationToken);
            if (vt.IsCompletedSuccessfully)
            {
                _pos = 0;
                return default;
            }
            return AwaitAndResetAsync(vt);
#endif
        }

#if NETSTANDARD2_0
        private async ValueTask FlushToAsyncCoreNs20(Stream stream, CancellationToken cancellationToken)
        {
            await stream.WriteAsync(_buffer, 0, _pos, cancellationToken).ConfigureAwait(false);
            _pos = 0;
        }
#else
        private async ValueTask AwaitAndResetAsync(ValueTask writeTask)
        {
            try
            {
                await writeTask.ConfigureAwait(false);
            }
            finally
            {
                _pos = 0;
            }
        }
#endif

        public int GetCommittedSize()
        {
            return _pos == 0 ? 256 : _pos;
        }

        public void Dispose()
        {
            if (_buffer.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_buffer);
                _buffer = Array.Empty<byte>();
            }
        }
    }
}
