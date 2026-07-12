using System;
using System.Buffers;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
#if NET6_0_OR_GREATER
using System.Runtime.Intrinsics.Arm;
#endif

namespace Magicodes.IE.IO
{

    internal sealed class ForwardOnlyZipWriter : IDisposable
    {
        private readonly Stream _output;
        private readonly bool _leaveOpen;
        private readonly List<CentralDirectoryEntry> _entries = new();
        private EntryWriteStream? _activeEntry;
        private long _position;
        private bool _disposed;

        public ForwardOnlyZipWriter(Stream output, bool leaveOpen = true)
        {
            _output = output ?? throw new ArgumentNullException(nameof(output));
            _leaveOpen = leaveOpen;
        }

        public Stream OpenEntry(string name, CompressionLevel compression)
        {
            if (name is null) throw new ArgumentNullException(nameof(name));
            EnsureNotDisposed();
            if (_activeEntry is not null)
                throw new InvalidOperationException("Previous zip entry has not been closed.");

            var utf8Name = RentUtf8Name(name);
            try
            {
                ushort method = compression == CompressionLevel.NoCompression ? (ushort)0 : (ushort)8;
                const ushort flags = 0x0808;
                DosDateTime dos = DosDateTime.From(DateTime.Now);
                uint offset = CheckedToUInt32(_position);

                WriteLocalHeader(utf8Name.Buffer, utf8Name.Length, method, flags, dos.Time, dos.Date);

                var entry = new CentralDirectoryEntry(utf8Name.Buffer, utf8Name.Length, method, flags, dos.Time, dos.Date, offset);
                _activeEntry = new EntryWriteStream(this, entry, compression);
                return _activeEntry;
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(utf8Name.Buffer);
                throw;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_activeEntry is not null)
                    _activeEntry.Dispose();

                WriteCentralDirectory();
            }
            finally
            {
                for (int i = 0; i < _entries.Count; i++)
                    _entries[i].ReturnNameBuffer();
                if (!_leaveOpen)
                    _output.Dispose();
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_activeEntry is not null)
                    await _activeEntry.DisposeAsync().ConfigureAwait(false);

                await WriteCentralDirectoryAsync(CancellationToken.None).ConfigureAwait(false);
            }
            finally
            {
                for (int i = 0; i < _entries.Count; i++)
                    _entries[i].ReturnNameBuffer();
                if (!_leaveOpen)
                {
#if NETSTANDARD2_0
                    _output.Dispose();
#else
                    await _output.DisposeAsync().ConfigureAwait(false);
#endif
                }
            }
        }

        internal static ValueTask DisposeEntryStreamAsync(Stream entryStream)
        {
#if NETSTANDARD2_0
            entryStream.Dispose();
            return default;
#else
            return entryStream.DisposeAsync();
#endif
        }

        internal void WriteRaw(byte[] buffer, int offset, int count)
        {
            if (count == 0) return;
            if (_position + count > uint.MaxValue) ThrowZip32LimitExceeded();
#if NETSTANDARD2_0
            _output.Write(buffer, offset, count);
#else
            _output.Write(buffer.AsSpan(offset, count));
#endif
            _position += count;
        }

        internal async Task WriteRawAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            if (count == 0) return;
            if (_position + count > uint.MaxValue) ThrowZip32LimitExceeded();
#if NETSTANDARD2_0
            await _output.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            _position += count;
#else
            await _output.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            _position += count;
#endif
        }

#if !NETSTANDARD2_0
        internal ValueTask WriteRawAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken)
        {
            if (buffer.Length == 0) return default;
            if (_position + buffer.Length > uint.MaxValue) ThrowZip32LimitExceeded();
            var vt = _output.WriteAsync(buffer, cancellationToken);
            if (vt.IsCompletedSuccessfully)
            {
                _position += buffer.Length;
                return default;
            }
            return AwaitAndUpdatePosition(vt, buffer.Length);
        }

        private async ValueTask AwaitAndUpdatePosition(ValueTask writeTask, int length)
        {
            await writeTask.ConfigureAwait(false);
            _position += length;
        }
#endif

#if !NETSTANDARD2_0
        internal void WriteRaw(ReadOnlySpan<byte> buffer)
        {
            if (buffer.Length == 0) return;
            _output.Write(buffer);
            _position += buffer.Length;
        }
#endif

        private void CompleteEntry(CentralDirectoryEntry entry, uint crc32, uint compressedSize, uint uncompressedSize)
        {
            WriteDataDescriptor(crc32, compressedSize, uncompressedSize);
            entry.Crc32 = crc32;
            entry.CompressedSize = compressedSize;
            entry.UncompressedSize = uncompressedSize;
            _entries.Add(entry);
            _activeEntry = null;
        }

        private async Task CompleteEntryAsync(CentralDirectoryEntry entry, uint crc32, uint compressedSize, uint uncompressedSize, CancellationToken cancellationToken)
        {
            await WriteDataDescriptorAsync(crc32, compressedSize, uncompressedSize, cancellationToken).ConfigureAwait(false);
            entry.Crc32 = crc32;
            entry.CompressedSize = compressedSize;
            entry.UncompressedSize = uncompressedSize;
            _entries.Add(entry);
            _activeEntry = null;
        }

        private void WriteLocalHeader(byte[] nameBytes, int nameLength, ushort method, ushort flags, ushort dosTime, ushort dosDate)
        {
            int len = 30 + nameLength;
            byte[] rented = ArrayPool<byte>.Shared.Rent(len);
            try
            {
                var span = rented.AsSpan(0, len);
                span.Clear();

                WriteUInt32(span, 0, 0x04034B50u);
                WriteUInt16(span, 4, 20);
                WriteUInt16(span, 6, flags);
                WriteUInt16(span, 8, method);
                WriteUInt16(span, 10, dosTime);
                WriteUInt16(span, 12, dosDate);
                WriteUInt32(span, 14, 0);
                WriteUInt32(span, 18, 0);
                WriteUInt32(span, 22, 0);
                WriteUInt16(span, 26, CheckedToUInt16(nameLength));
                WriteUInt16(span, 28, 0);
                nameBytes.AsSpan(0, nameLength).CopyTo(span.Slice(30));

#if NETSTANDARD2_0
                WriteRaw(rented, 0, len);
#else
                WriteRaw(span);
#endif
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private void WriteDataDescriptor(uint crc32, uint compressedSize, uint uncompressedSize)
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(16);
            try
            {
                var span = rented.AsSpan(0, 16);
                WriteUInt32(span, 0, 0x08074B50u);
                WriteUInt32(span, 4, crc32);
                WriteUInt32(span, 8, compressedSize);
                WriteUInt32(span, 12, uncompressedSize);
#if NETSTANDARD2_0
                WriteRaw(rented, 0, 16);
#else
                WriteRaw(span);
#endif
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private async Task WriteDataDescriptorAsync(uint crc32, uint compressedSize, uint uncompressedSize, CancellationToken cancellationToken)
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(16);
            try
            {
                var span = rented.AsSpan(0, 16);
                WriteUInt32(span, 0, 0x08074B50u);
                WriteUInt32(span, 4, crc32);
                WriteUInt32(span, 8, compressedSize);
                WriteUInt32(span, 12, uncompressedSize);
                await WriteRawAsync(rented, 0, 16, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private void WriteCentralDirectory()
        {
            long centralDirectoryOffset = _position;

            for (int i = 0; i < _entries.Count; i++)
            {
                WriteCentralDirectoryHeader(_entries[i]);
            }

            long centralDirectorySize = _position - centralDirectoryOffset;
            WriteEndOfCentralDirectory(CheckedToUInt16(_entries.Count), CheckedToUInt32(centralDirectorySize), CheckedToUInt32(centralDirectoryOffset));
        }

        private async Task WriteCentralDirectoryAsync(CancellationToken cancellationToken)
        {
            long centralDirectoryOffset = _position;

            for (int i = 0; i < _entries.Count; i++)
            {
                await WriteCentralDirectoryHeaderAsync(_entries[i], cancellationToken).ConfigureAwait(false);
            }

            long centralDirectorySize = _position - centralDirectoryOffset;
            await WriteEndOfCentralDirectoryAsync(CheckedToUInt16(_entries.Count), CheckedToUInt32(centralDirectorySize), CheckedToUInt32(centralDirectoryOffset), cancellationToken).ConfigureAwait(false);
        }

        private void WriteCentralDirectoryHeader(CentralDirectoryEntry entry)
        {
            int len = 46 + entry.NameLength;
            byte[] rented = ArrayPool<byte>.Shared.Rent(len);
            try
            {
                var span = rented.AsSpan(0, len);
                span.Clear();

                WriteUInt32(span, 0, 0x02014B50u);
                WriteUInt16(span, 4, 20);
                WriteUInt16(span, 6, 20);
                WriteUInt16(span, 8, entry.Flags);
                WriteUInt16(span, 10, entry.Method);
                WriteUInt16(span, 12, entry.DosTime);
                WriteUInt16(span, 14, entry.DosDate);
                WriteUInt32(span, 16, entry.Crc32);
                WriteUInt32(span, 20, entry.CompressedSize);
                WriteUInt32(span, 24, entry.UncompressedSize);
                WriteUInt16(span, 28, CheckedToUInt16(entry.NameLength));
                WriteUInt16(span, 30, 0);
                WriteUInt16(span, 32, 0);
                WriteUInt16(span, 34, 0);
                WriteUInt16(span, 36, 0);
                WriteUInt32(span, 38, 0);
                WriteUInt32(span, 42, entry.LocalHeaderOffset);
                entry.NameBuffer.AsSpan(0, entry.NameLength).CopyTo(span.Slice(46));

#if NETSTANDARD2_0
                WriteRaw(rented, 0, len);
#else
                WriteRaw(span);
#endif
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private async Task WriteCentralDirectoryHeaderAsync(CentralDirectoryEntry entry, CancellationToken cancellationToken)
        {
            int len = 46 + entry.NameLength;
            byte[] rented = ArrayPool<byte>.Shared.Rent(len);
            try
            {
                var span = rented.AsSpan(0, len);
                span.Clear();

                WriteUInt32(span, 0, 0x02014B50u);
                WriteUInt16(span, 4, 20);
                WriteUInt16(span, 6, 20);
                WriteUInt16(span, 8, entry.Flags);
                WriteUInt16(span, 10, entry.Method);
                WriteUInt16(span, 12, entry.DosTime);
                WriteUInt16(span, 14, entry.DosDate);
                WriteUInt32(span, 16, entry.Crc32);
                WriteUInt32(span, 20, entry.CompressedSize);
                WriteUInt32(span, 24, entry.UncompressedSize);
                WriteUInt16(span, 28, CheckedToUInt16(entry.NameLength));
                WriteUInt16(span, 30, 0);
                WriteUInt16(span, 32, 0);
                WriteUInt16(span, 34, 0);
                WriteUInt16(span, 36, 0);
                WriteUInt32(span, 38, 0);
                WriteUInt32(span, 42, entry.LocalHeaderOffset);
                entry.NameBuffer.AsSpan(0, entry.NameLength).CopyTo(span.Slice(46));

                await WriteRawAsync(rented, 0, len, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private void WriteEndOfCentralDirectory(ushort entryCount, uint centralDirectorySize, uint centralDirectoryOffset)
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(22);
            try
            {
                var span = rented.AsSpan(0, 22);
                span.Clear();

                WriteUInt32(span, 0, 0x06054B50u);
                WriteUInt16(span, 4, 0);
                WriteUInt16(span, 6, 0);
                WriteUInt16(span, 8, entryCount);
                WriteUInt16(span, 10, entryCount);
                WriteUInt32(span, 12, centralDirectorySize);
                WriteUInt32(span, 16, centralDirectoryOffset);
                WriteUInt16(span, 20, 0);

#if NETSTANDARD2_0
                WriteRaw(rented, 0, 22);
#else
                WriteRaw(span);
#endif
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private async Task WriteEndOfCentralDirectoryAsync(ushort entryCount, uint centralDirectorySize, uint centralDirectoryOffset, CancellationToken cancellationToken)
        {
            byte[] rented = ArrayPool<byte>.Shared.Rent(22);
            try
            {
                var span = rented.AsSpan(0, 22);
                span.Clear();

                WriteUInt32(span, 0, 0x06054B50u);
                WriteUInt16(span, 4, 0);
                WriteUInt16(span, 6, 0);
                WriteUInt16(span, 8, entryCount);
                WriteUInt16(span, 10, entryCount);
                WriteUInt32(span, 12, centralDirectorySize);
                WriteUInt32(span, 16, centralDirectoryOffset);
                WriteUInt16(span, 20, 0);

                await WriteRawAsync(rented, 0, 22, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(rented);
            }
        }

        private static void WriteUInt16(Span<byte> dest, int offset, ushort value)
        {
            dest[offset] = (byte)value;
            dest[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteUInt32(Span<byte> dest, int offset, uint value)
        {
            dest[offset] = (byte)value;
            dest[offset + 1] = (byte)(value >> 8);
            dest[offset + 2] = (byte)(value >> 16);
            dest[offset + 3] = (byte)(value >> 24);
        }

        private static ushort CheckedToUInt16(int value)
        {
            if ((uint)value > ushort.MaxValue) throw new NotSupportedException("Zip field exceeds UInt16 range.");
            return (ushort)value;
        }

        private static Utf8NameBuffer RentUtf8Name(string name)
        {
#if NETSTANDARD2_0
            int byteCount = Encoding.UTF8.GetByteCount(name);
            if (byteCount > ushort.MaxValue)
                throw new ArgumentException("Entry name is too long.", nameof(name));

            byte[] rented = ArrayPool<byte>.Shared.Rent(byteCount);
            int written = Encoding.UTF8.GetBytes(name, 0, name.Length, rented, 0);
            return new Utf8NameBuffer(rented, written);
#else
            int byteCount = Encoding.UTF8.GetByteCount(name.AsSpan());
            if (byteCount > ushort.MaxValue)
                throw new ArgumentException("Entry name is too long.", nameof(name));

            byte[] rented = ArrayPool<byte>.Shared.Rent(byteCount);
            int written = Encoding.UTF8.GetBytes(name.AsSpan(), rented.AsSpan());
            return new Utf8NameBuffer(rented, written);
#endif
        }

        internal static uint CheckedToUInt32(long value)
        {
            if ((ulong)value > uint.MaxValue) ThrowZip32LimitExceeded();
            return (uint)value;
        }

        private static void ThrowZip32LimitExceeded()
        {
            throw new NotSupportedException(
                "The output exceeds the 4 GB (ZIP32) size limit supported by this forward-only zip writer. " +
                "Magicodes.IE.IO does not emit ZIP64 archives. To write workbooks larger than 4 GB, " +
                "split the data across multiple files or sheets, or write to a consumer that supports ZIP64.");
        }

        private void EnsureNotDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ForwardOnlyZipWriter));
        }

        internal interface ICompressor : IDisposable
        {
            void Write(ReadOnlySpan<byte> data);
            void Flush();
            void Finish();
            ValueTask WriteAsync(ReadOnlyMemory<byte> data, CancellationToken cancellationToken);
            ValueTask FlushAsync();
            ValueTask FinishAsync();
        }

        internal static Func<Stream, CompressionLevel, ICompressor>? CompressorProvider;

        private static ICompressor CreateDefaultCompressor(Stream sink, CompressionLevel compression)
        {
            return compression == CompressionLevel.NoCompression
                ? (ICompressor)new CopyCompressor(sink)
                : new DeflateStreamCompressor(sink, compression);
        }

        private sealed class DeflateStreamCompressor : ICompressor
        {
            private readonly DeflateStream _deflate;

            public DeflateStreamCompressor(Stream sink, CompressionLevel level)
            {
                _deflate = new DeflateStream(sink, level, leaveOpen: true);
            }

            public void Write(ReadOnlySpan<byte> data)
            {
                if (data.Length == 0) return;
#if NETSTANDARD2_0
                byte[] tmp = data.ToArray();
                _deflate.Write(tmp, 0, tmp.Length);
#else
                _deflate.Write(data);
#endif
            }

            public void Flush() => _deflate.Flush();

            public void Finish() => _deflate.Dispose();

            public void Dispose() => _deflate.Dispose();

            public ValueTask WriteAsync(ReadOnlyMemory<byte> data, CancellationToken cancellationToken)
            {
                if (data.Length == 0) return default;
#if NETSTANDARD2_0
                var tmp = data.ToArray();
                return new ValueTask(_deflate.WriteAsync(tmp, 0, tmp.Length, cancellationToken));
#else
                return _deflate.WriteAsync(data, cancellationToken);
#endif
            }

            public ValueTask FlushAsync()
            {
#if NETSTANDARD2_0
                _deflate.Flush();
                return default;
#else
                return new ValueTask(_deflate.FlushAsync(CancellationToken.None));
#endif
            }

            public ValueTask FinishAsync()
            {
#if NETSTANDARD2_0
                _deflate.Dispose();
                return default;
#else
                return _deflate.DisposeAsync();
#endif
            }
        }

        private sealed class CopyCompressor : ICompressor
        {
            private readonly Stream _sink;

            public CopyCompressor(Stream sink) => _sink = sink;

            public void Write(ReadOnlySpan<byte> data)
            {
                if (data.Length == 0) return;
#if NETSTANDARD2_0
                byte[] tmp = data.ToArray();
                _sink.Write(tmp, 0, tmp.Length);
#else
                _sink.Write(data);
#endif
            }

            public ValueTask WriteAsync(ReadOnlyMemory<byte> data, CancellationToken cancellationToken)
            {
                if (data.Length == 0) return default;
#if NETSTANDARD2_0
                var tmp = data.ToArray();
                return new ValueTask(_sink.WriteAsync(tmp, 0, tmp.Length, cancellationToken));
#else
                return _sink.WriteAsync(data, cancellationToken);
#endif
            }

            public ValueTask FlushAsync() => default;

            public ValueTask FinishAsync() => default;

            public void Flush() => _sink.Flush();

            public void Finish() { }

            public void Dispose() { }
        }

        private sealed class EntryWriteStream : Stream
        {
            private readonly ForwardOnlyZipWriter _owner;
            private readonly CentralDirectoryEntry _entry;
            private readonly CountingWriteStream _countingStream;
            private readonly ICompressor _compressor;
            private readonly ICrc32 _crc32;
            private long _uncompressedSize;
            private bool _disposed;

            public EntryWriteStream(ForwardOnlyZipWriter owner, CentralDirectoryEntry entry, CompressionLevel compression)
            {
                _owner = owner;
                _entry = entry;
                _countingStream = new CountingWriteStream(owner, bufferWrites: false);
                _compressor = CompressorProvider is not null
                    ? CompressorProvider(_countingStream, compression)
                    : CreateDefaultCompressor(_countingStream, compression);
                _crc32 = CreateDefaultCrc32();
            }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => !_disposed;
            public override long Length => _uncompressedSize;

            public override long Position
            {
                get => _uncompressedSize;
                set => throw new NotSupportedException();
            }

            public override void Flush() => _compressor.Flush();

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count)
            {
                if (buffer is null) throw new ArgumentNullException(nameof(buffer));
                if ((uint)offset > (uint)buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));
                if ((uint)count > (uint)(buffer.Length - offset)) throw new ArgumentOutOfRangeException(nameof(count));
                if (count == 0) return;

                EnsureNotDisposed();
                _crc32.Update(buffer.AsSpan(offset, count));
                _uncompressedSize += count;
                _compressor.Write(buffer.AsSpan(offset, count));
            }

            public override Task WriteAsync(byte[] buffer, int offset, int count, System.Threading.CancellationToken cancellationToken)
            {
                if (cancellationToken.IsCancellationRequested)
                    return Task.FromCanceled(cancellationToken);

                if (buffer is null) throw new ArgumentNullException(nameof(buffer));
                if ((uint)offset > (uint)buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));
                if ((uint)count > (uint)(buffer.Length - offset)) throw new ArgumentOutOfRangeException(nameof(count));
                if (count == 0) return Task.CompletedTask;

                EnsureNotDisposed();
                _crc32.Update(buffer.AsSpan(offset, count));
                _uncompressedSize += count;
                return _compressor.WriteAsync(new ReadOnlyMemory<byte>(buffer, offset, count), cancellationToken).AsTask();
            }

#if !NETSTANDARD2_0
            public override void Write(ReadOnlySpan<byte> buffer)
            {
                if (buffer.Length == 0) return;

                EnsureNotDisposed();
                _crc32.Update(buffer);
                _uncompressedSize += buffer.Length;
                _compressor.Write(buffer);
            }

            public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, System.Threading.CancellationToken cancellationToken = default)
            {
                if (cancellationToken.IsCancellationRequested)
                    return ValueTask.FromCanceled(cancellationToken);

                if (buffer.Length == 0) return default;

                EnsureNotDisposed();
                _crc32.Update(buffer.Span);
                _uncompressedSize += buffer.Length;
                return _compressor.WriteAsync(buffer, cancellationToken);
            }
#endif

            public override Task FlushAsync(System.Threading.CancellationToken cancellationToken)
            {
                EnsureNotDisposed();
                return _compressor.FlushAsync().AsTask();
            }

#if NETSTANDARD2_0
            public ValueTask DisposeAsync()
#else
            public override ValueTask DisposeAsync()
#endif
            {
                if (_disposed) return default;
                _disposed = true;
                return DisposeAsyncCore();
            }

            private async ValueTask DisposeAsyncCore()
            {
                try
                {
                    await _compressor.FinishAsync().ConfigureAwait(false);
                }
                finally
                {
                    try
                    {
                        await _countingStream.DisposeAsync().ConfigureAwait(false);
                    }
                    finally
                    {
                        uint compressedSize = CheckedToUInt32(_countingStream.BytesWritten);
                        uint uncompressedSize = CheckedToUInt32(_uncompressedSize);
                        await _owner.CompleteEntryAsync(_entry, _crc32.Result, compressedSize, uncompressedSize, System.Threading.CancellationToken.None).ConfigureAwait(false);
                    }
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (_disposed) return;
                _disposed = true;

                try
                {
                    _compressor.Finish();
                }
                finally
                {
                    try
                    {
                        _countingStream.Dispose();
                    }
                    finally
                    {
                        uint compressedSize = CheckedToUInt32(_countingStream.BytesWritten);
                        uint uncompressedSize = CheckedToUInt32(_uncompressedSize);
                        _owner.CompleteEntry(_entry, _crc32.Result, compressedSize, uncompressedSize);

                        base.Dispose(disposing);
                    }
                }
            }

            private void EnsureNotDisposed()
            {
                if (_disposed) throw new ObjectDisposedException(nameof(EntryWriteStream));
            }
        }

        private sealed class CountingWriteStream : Stream
        {
            private readonly ForwardOnlyZipWriter _owner;
            private readonly byte[]? _buffer;
            private int _buffered;
            private bool _disposed;

            public CountingWriteStream(ForwardOnlyZipWriter owner, bool bufferWrites)
            {
                _owner = owner;
                if (bufferWrites)
                    _buffer = ArrayPool<byte>.Shared.Rent(16 * 1024);
            }

            public long BytesWritten { get; private set; }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => BytesWritten;

            public override long Position
            {
                get => BytesWritten;
                set => throw new NotSupportedException();
            }

            public override void Flush()
            {
                FlushBuffered();
            }

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count)
            {
                if (buffer is null) throw new ArgumentNullException(nameof(buffer));
                if ((uint)offset > (uint)buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));
                if ((uint)count > (uint)(buffer.Length - offset)) throw new ArgumentOutOfRangeException(nameof(count));
                if (count == 0) return;

                EnsureNotDisposed();
                WriteBuffered(buffer.AsSpan(offset, count));
                BytesWritten += count;
            }

#if !NETSTANDARD2_0
            public override void Write(ReadOnlySpan<byte> buffer)
            {
                if (buffer.Length == 0) return;

                EnsureNotDisposed();
                WriteBuffered(buffer);
                BytesWritten += buffer.Length;
            }
#endif

            public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                if (buffer is null) throw new ArgumentNullException(nameof(buffer));
                if ((uint)offset > (uint)buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));
                if ((uint)count > (uint)(buffer.Length - offset)) throw new ArgumentOutOfRangeException(nameof(count));
                if (count == 0) return Task.CompletedTask;
                EnsureNotDisposed();
                return WriteBufferedAsync(buffer, offset, count, cancellationToken);
            }

#if !NETSTANDARD2_0
            public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                if (buffer.Length == 0) return default;
                EnsureNotDisposed();
                BytesWritten += buffer.Length;
                if (_buffer is null)
                {
                    return _owner.WriteRawAsync(buffer, cancellationToken);
                }
                WriteBuffered(buffer.Span);
                return default;
            }
#endif

            private Task WriteBufferedAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                if (count == 0) return Task.CompletedTask;
                BytesWritten += count;
                if (_buffer is null)
                {
                    return _owner.WriteRawAsync(buffer, offset, count, cancellationToken);
                }
                WriteBuffered(buffer.AsSpan(offset, count));
                return Task.CompletedTask;
            }

            public override Task FlushAsync(CancellationToken cancellationToken)
            {
                return FlushBufferedAsync(cancellationToken);
            }

            private Task FlushBufferedAsync(CancellationToken cancellationToken)
            {
                if (_buffer is null || _buffered == 0) return Task.CompletedTask;
                var t = _owner.WriteRawAsync(_buffer, 0, _buffered, cancellationToken);
                _buffered = 0;
                return t;
            }

#if NETSTANDARD2_0
            public ValueTask DisposeAsync()
#else
            public override ValueTask DisposeAsync()
#endif
            {
                if (_disposed) return default;
                _disposed = true;
                FlushBuffered();
                if (_buffer is not null)
                    ArrayPool<byte>.Shared.Return(_buffer);
#if NETSTANDARD2_0
                return default;
#else
                return base.DisposeAsync();
#endif
            }

            protected override void Dispose(bool disposing)
            {
                if (_disposed) return;
                _disposed = true;
                FlushBuffered();
                if (_buffer is not null)
                    ArrayPool<byte>.Shared.Return(_buffer);
                base.Dispose(disposing);
            }

            private void FlushBuffered()
            {
                if (_buffer is null || _buffered == 0) return;
                _owner.WriteRaw(_buffer, 0, _buffered);
                _buffered = 0;
            }

            private void WriteBuffered(ReadOnlySpan<byte> buffer)
            {
                if (_buffer is null)
                {
#if NETSTANDARD2_0
                    var tmp = buffer.ToArray();
                    _owner.WriteRaw(tmp, 0, tmp.Length);
#else
                    _owner.WriteRaw(buffer);
#endif
                    return;
                }

                if (buffer.Length >= _buffer.Length)
                {
                    FlushBuffered();
#if NETSTANDARD2_0
                    var tmp = buffer.ToArray();
                    _owner.WriteRaw(tmp, 0, tmp.Length);
#else
                    _owner.WriteRaw(buffer);
#endif
                    return;
                }

                if (_buffered + buffer.Length > _buffer.Length)
                    FlushBuffered();

                buffer.CopyTo(_buffer.AsSpan(_buffered));
                _buffered += buffer.Length;
            }

            private void EnsureNotDisposed()
            {
                if (_disposed) throw new ObjectDisposedException(nameof(CountingWriteStream));
            }
        }

        private sealed class CentralDirectoryEntry
        {
            public CentralDirectoryEntry(byte[] nameBuffer, int nameLength, ushort method, ushort flags, ushort dosTime, ushort dosDate, uint localHeaderOffset)
            {
                NameBuffer = nameBuffer;
                NameLength = nameLength;
                Method = method;
                Flags = flags;
                DosTime = dosTime;
                DosDate = dosDate;
                LocalHeaderOffset = localHeaderOffset;
            }

            public byte[] NameBuffer { get; private set; }
            public int NameLength { get; }
            public ushort Method { get; }
            public ushort Flags { get; }
            public ushort DosTime { get; }
            public ushort DosDate { get; }
            public uint LocalHeaderOffset { get; }
            public uint Crc32 { get; set; }
            public uint CompressedSize { get; set; }
            public uint UncompressedSize { get; set; }

            public void ReturnNameBuffer()
            {
                if (NameBuffer.Length == 0) return;
                ArrayPool<byte>.Shared.Return(NameBuffer);
                NameBuffer = Array.Empty<byte>();
            }
        }

        private readonly struct Utf8NameBuffer
        {
            public Utf8NameBuffer(byte[] buffer, int length)
            {
                Buffer = buffer;
                Length = length;
            }

            public byte[] Buffer { get; }
            public int Length { get; }
        }

        private readonly struct DosDateTime
        {
            public DosDateTime(ushort time, ushort date)
            {
                Time = time;
                Date = date;
            }

            public ushort Time { get; }
            public ushort Date { get; }

            public static DosDateTime From(DateTime value)
            {
                if (value.Kind == DateTimeKind.Unspecified) value = DateTime.SpecifyKind(value, DateTimeKind.Local);
                else if (value.Kind == DateTimeKind.Utc)
                    throw new ArgumentException("DosDateTime.From requires local time. Convert UTC to local time before calling.", nameof(value));
                if (value.Year < 1980) value = new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Local);

                ushort time = (ushort)((value.Hour << 11) | (value.Minute << 5) | (value.Second >> 1));
                ushort date = (ushort)(((value.Year - 1980) << 9) | (value.Month << 5) | value.Day);
                return new DosDateTime(time, date);
            }
        }

        private static readonly uint[] Crc32Table = BuildCrc32Table();

        // Self-check the CRC-32 implementation at type load. This guards the SIMD/intrinsic path
        // (active on ARM64 in Release builds) so a wrong CRC never silently corrupts every xlsx.
        static ForwardOnlyZipWriter()
        {
            var data = System.Text.Encoding.ASCII.GetBytes("123456789");
            const uint expected = 0xCBF43926u;

            var intrinsic = new IntrinsicCrc32();
            intrinsic.Update(data);
            if (intrinsic.Result != expected)
                throw new InvalidOperationException($"IntrinsicCrc32 self-test failed: {intrinsic.Result:X8}");

            var sliced = new SlicingBy8Crc32();
            sliced.Update(data);
            if (sliced.Result != expected)
                throw new InvalidOperationException($"SlicingBy8Crc32 self-test failed: {sliced.Result:X8}");
        }

        internal interface ICrc32
        {
            void Update(ReadOnlySpan<byte> data);
            uint Result { get; }
        }

        internal static Func<ICrc32>? Crc32Provider;

        private static ICrc32 CreateDefaultCrc32()
        {
            return Crc32Provider is not null
                ? Crc32Provider()
#if NET6_0_OR_GREATER
                : (Crc32.IsSupported
                    ? (ICrc32)new IntrinsicCrc32()
                    : new SlicingBy8Crc32());
#else
                : new SlicingBy8Crc32();
#endif
        }

        private sealed class IntrinsicCrc32 : ICrc32
        {
            private uint _crc = 0xFFFFFFFFu;

            public void Update(ReadOnlySpan<byte> data)
            {
#if NET6_0_OR_GREATER
                if (Crc32.IsSupported)
                {
                    int i = 0, len = data.Length;
                    while (i + 8 <= len)
                    {
                        ulong v = BinaryPrimitives.ReadUInt64LittleEndian(data.Slice(i));
                        _crc = Crc32.ComputeCrc32(_crc, (uint)v);
                        _crc = Crc32.ComputeCrc32(_crc, (uint)(v >> 32));
                        i += 8;
                    }
                    if (i + 4 <= len)
                    {
                        _crc = Crc32.ComputeCrc32(_crc, BinaryPrimitives.ReadUInt32LittleEndian(data.Slice(i)));
                        i += 4;
                    }
                    if (i + 2 <= len)
                    {
                        _crc = Crc32.ComputeCrc32(_crc, BinaryPrimitives.ReadUInt16LittleEndian(data.Slice(i)));
                        i += 2;
                    }
                    for (; i < len; i++)
                        _crc = Crc32.ComputeCrc32(_crc, data[i]);
                    return;
                }
#endif
                for (int i = 0; i < data.Length; i++)
                    _crc = Crc32Table[(_crc ^ data[i]) & 0xFF] ^ (_crc >> 8);
            }

            public uint Result => ~_crc;
        }

        private sealed class SlicingBy8Crc32 : ICrc32
        {
            private uint _crc = 0xFFFFFFFFu;

            public void Update(ReadOnlySpan<byte> data)
            {
                uint c = _crc;
                int len = data.Length;
                int i = 0;
                while (i + 8 <= len)
                {
                    c ^= BinaryPrimitives.ReadUInt32LittleEndian(data.Slice(i));
                    c = Crc32Table[0x700 + (c & 0xFF)] ^
                        Crc32Table[0x600 + ((c >> 8) & 0xFF)] ^
                        Crc32Table[0x500 + ((c >> 16) & 0xFF)] ^
                        Crc32Table[0x400 + (c >> 24)] ^
                        Crc32Table[0x300 + data[i + 4]] ^
                        Crc32Table[0x200 + data[i + 5]] ^
                        Crc32Table[0x100 + data[i + 6]] ^
                        Crc32Table[0x000 + data[i + 7]];
                    i += 8;
                }
                for (; i < len; i++)
                    c = Crc32Table[0x000 + (byte)(c ^ data[i])] ^ (c >> 8);
                _crc = c;
            }

            public uint Result => ~_crc;
        }

        private static uint[] BuildCrc32Table()
        {
            var table = new uint[8 * 256];
            for (int n = 0; n < 256; n++)
            {
                uint value = (uint)n;
                for (int bit = 0; bit < 8; bit++)
                    value = (value & 1) != 0 ? 0xEDB88320u ^ (value >> 1) : value >> 1;
                table[n] = value;
            }
            for (int n = 0; n < 256; n++)
            {
                uint value = table[n];
                for (int k = 1; k < 8; k++)
                {
                    value = table[value & 0xFF] ^ (value >> 8);
                    table[k * 256 + n] = value;
                }
            }
            return table;
        }
    }
}
