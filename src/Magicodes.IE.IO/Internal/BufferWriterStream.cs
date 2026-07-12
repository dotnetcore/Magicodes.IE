using System;
using System.Buffers;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Magicodes.IE.IO
{

    internal sealed class BufferWriterStream : Stream
    {
        private readonly IBufferWriter<byte> _writer;
        private bool _disposed;

        public BufferWriterStream(IBufferWriter<byte> writer)
        {
            _writer = writer ?? throw new ArgumentNullException(nameof(writer));
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => !_disposed;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
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
            var dest = _writer.GetSpan(count);
            buffer.AsSpan(offset, count).CopyTo(dest);
            _writer.Advance(count);
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            if (cancellationToken.IsCancellationRequested)
                return Task.FromCanceled(cancellationToken);

            Write(buffer, offset, count);
            return Task.CompletedTask;
        }

#if !NETSTANDARD2_0
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            if (buffer.Length == 0) return;

            EnsureNotDisposed();
            var dest = _writer.GetSpan(buffer.Length);
            buffer.CopyTo(dest);
            _writer.Advance(buffer.Length);
        }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (cancellationToken.IsCancellationRequested)
                return ValueTask.FromCanceled(cancellationToken);

            Write(buffer.Span);
            return ValueTask.CompletedTask;
        }
#endif

        protected override void Dispose(bool disposing)
        {
            _disposed = true;
            base.Dispose(disposing);
        }

        private void EnsureNotDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(BufferWriterStream));
        }
    }
}
