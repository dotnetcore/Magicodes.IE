// net471 does not ship ArrayBufferWriter<T> (it was added in .NET Core 3.0 / netstandard2.1).
// The BCL IBufferWriter<T> interface IS available via the System.Memory package, so we provide a
// minimal ArrayBufferWriter<T> implementation for the .NET Framework target only. On .NET Core /
// .NET 5+ the real System.Buffers.ArrayBufferWriter<T> is used instead.
#if NETFRAMEWORK
namespace System.Buffers
{
    using System;

    internal sealed class ArrayBufferWriter<T> : IBufferWriter<T>
    {
        private T[] _buffer = Array.Empty<T>();
        private int _count;

        public ArrayBufferWriter()
        {
        }

        public ArrayBufferWriter(int initialCapacity)
        {
            if (initialCapacity > 0)
                _buffer = new T[initialCapacity];
        }

        public ReadOnlyMemory<T> WrittenMemory => _buffer.AsMemory(0, _count);

        public ReadOnlySpan<T> WrittenSpan => _buffer.AsSpan(0, _count);

        public int WrittenCount => _count;

        public void Clear() => _count = 0;

        public void Write(ReadOnlySpan<T> data)
        {
            if (_count + data.Length > _buffer.Length)
                Resize(_count + data.Length);
            data.CopyTo(_buffer.AsSpan(_count));
            _count += data.Length;
        }

        public void Write(T[] data, int index, int count) => Write(data.AsSpan(index, count));

        public void Advance(int count) => _count += count;

        public Memory<T> GetMemory(int sizeHint = 0)
        {
            EnsureCapacity(sizeHint);
            return _buffer.AsMemory(_count);
        }

        public Span<T> GetSpan(int sizeHint = 0)
        {
            EnsureCapacity(sizeHint);
            return _buffer.AsSpan(_count);
        }

        private void EnsureCapacity(int sizeHint)
        {
            int needed = _count + Math.Max(sizeHint, 1);
            if (needed > _buffer.Length)
                Resize(needed);
        }

        private void Resize(int minimum)
        {
            int newSize = _buffer.Length == 0 ? 256 : _buffer.Length * 2;
            if (newSize < minimum)
                newSize = minimum;
            Array.Resize(ref _buffer, newSize);
        }
    }
}
#endif
