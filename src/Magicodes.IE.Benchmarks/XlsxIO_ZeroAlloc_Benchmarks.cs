using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Order;
using Magicodes.IE.IO;


namespace Magicodes.IE.Benchmarks
{
    [MemoryDiagnoser]
    [Orderer(SummaryOrderPolicy.FastestToSlowest)]
    [CategoriesColumn]
    public class XlsxIO_ZeroAlloc_Benchmarks
    {
        [XlsxExportable]
        public record ZA4(string C1, string C2, string C3, string C4);
        [XlsxExportable]
        public record Nullable4(string C1, int? C2, DateTime? C3, bool? C4);

        private ZA4[] _10k = null!;
        private ZA4[] _100k = null!;
        private FixedBufferWriter _bufferWriter10k = null!;
        private FixedBufferWriter _bufferWriter100k = null!;
        private FixedWriteStream _stream100k = null!;
        private byte[] _deflateInput = null!;

        [GlobalSetup]
        public void Setup()
        {
            _10k = Build(10_000);
            _100k = Build(100_000);
            _bufferWriter10k = new FixedBufferWriter(2 * 1024 * 1024);
            _bufferWriter100k = new FixedBufferWriter(8 * 1024 * 1024);
            _stream100k = new FixedWriteStream(8 * 1024 * 1024);
            // 诊断:可压缩的 sheet-XML 风格字节,供 DeflateStream sync vs async micro-benchmark。
            _deflateInput = Encoding.UTF8.GetBytes(string.Concat(Enumerable.Repeat(
                "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>s0</t></is></c></row>\r\n", 50_000)));
        }

        [BenchmarkCategory("convenience-10k"), Benchmark]
        public int Mio_10k_Bytes()
            => Xlsx.ToBytes(_10k).Length;

        [BenchmarkCategory("zeroalloc-10k"), Benchmark]
        public int Mio_10k_Write_BufferWriter()
        {
            _bufferWriter10k.Reset();
            Xlsx.Write(_bufferWriter10k, _10k);
            return _bufferWriter10k.WrittenCount;
        }

        [BenchmarkCategory("convenience-100k"), Benchmark]
        public int Mio_100k_Bytes()
            => Xlsx.ToBytes(_100k).Length;

        [BenchmarkCategory("zeroalloc-100k"), Benchmark]
        public int Mio_100k_Write_BufferWriter()
        {
            _bufferWriter100k.Reset();
            Xlsx.Write(_bufferWriter100k, _100k);
            return _bufferWriter100k.WrittenCount;
        }

        [BenchmarkCategory("zeroalloc-100k"), Benchmark]
        public long Mio_100k_Write_Stream_NoSeek()
        {
            _stream100k.Reset();
            using (var output = new NonSeekableWriteStream(_stream100k))
            {
                Xlsx.Write(output, _100k);
            }
            return _stream100k.Length;
        }

        // 异步路径零分配验证:走 WriteRowsAsync + CompleteAsync + DisposeAsync 全链路。
        // FixedWriteStream 的 WriteAsync 同步完成(不切线程),BDN 的 GetAllocatedBytesForCurrentThread
        // 能测准异步状态机 + I/O 链路上的真实分配(状态机 box 不可避免,但 ToArray/AsTask 应已消除)。
        [BenchmarkCategory("zeroalloc-100k"), Benchmark]
        public async Task<long> Mio_100k_WriteAsync_Stream()
        {
            _stream100k.Reset();
            await Xlsx.WriteAsync(_stream100k, ToAsyncEnum(_100k)).ConfigureAwait(false);
            return _stream100k.Length;
        }

        // 诊断:10k 行异步,与 100k 线性对比定位分配来源(固定开销 vs 每行开销)。
        [BenchmarkCategory("zeroalloc-10k"), Benchmark]
        public async Task<long> Mio_10k_WriteAsync_Stream()
        {
            _stream100k.Reset();
            await Xlsx.WriteAsync(_stream100k, ToAsyncEnum(_10k)).ConfigureAwait(false);
            return _stream100k.Length;
        }

        // IList 快路径:数据已物化(数组),同步遍历 + 异步 flush,无 IAsyncEnumerable 枚举层。
        [BenchmarkCategory("zeroalloc-100k"), Benchmark]
        public async Task<long> Mio_100k_WriteAsync_IList()
        {
            _stream100k.Reset();
            await Xlsx.WriteAsync(_stream100k, _100k).ConfigureAwait(false);
            return _stream100k.Length;
        }

        // ---- 真异步 I/O 场景(AsyncYieldingWriteStream,每次 WriteAsync 都 Yield) ----
        // 验证零分配改造 / 批量化 / IList 快路径在异步完成下的实际收益(同步完成 benchmark 测不出)。

        [BenchmarkCategory("asyncyield-100k"), Benchmark]
        public async Task<long> Mio_100k_AsyncYield_IList()
        {
            _stream100k.Reset();
            using var ys = new AsyncYieldingWriteStream(_stream100k);
            await Xlsx.WriteAsync(ys, _100k).ConfigureAwait(false);
            return _stream100k.Length;
        }

        [BenchmarkCategory("asyncyield-100k"), Benchmark]
        public async Task<long> Mio_100k_AsyncYield_AsyncEnum()
        {
            _stream100k.Reset();
            using var ys = new AsyncYieldingWriteStream(_stream100k);
            await Xlsx.WriteAsync(ys, ToAsyncEnum(_100k)).ConfigureAwait(false);
            return _stream100k.Length;
        }


        // 同步迭代器(无 Task.Yield):避免切线程,让 BDN 测准当前线程分配。
        private static async IAsyncEnumerable<T> ToAsyncEnum<T>(IEnumerable<T> source, [EnumeratorCancellation] CancellationToken ct = default)
        {
            using var e = source.GetEnumerator();
            while (e.MoveNext())
            {
                ct.ThrowIfCancellationRequested();
                yield return e.Current;
            }
        }

        private static ZA4[] Build(int n)
            => Enumerable.Range(0, n)
                .Select(i => new ZA4("s" + i, "c" + (i % 100), "p" + (i % 50), "t" + (i % 5)))
                .ToArray();

        // 诊断:0 行异步,隔离固定 async 开销(CompleteAsync/DisposeAsync/方法链)vs 每行开销。
        [BenchmarkCategory("deflate-diag"), Benchmark]
        public async Task<long> Mio_0_WriteAsync_IList()
        {
            _stream100k.Reset();
            await Xlsx.WriteAsync(_stream100k, Array.Empty<ZA4>()).ConfigureAwait(false);
            return _stream100k.Length;
        }

        // 诊断:同步 Write 包在 async Task 里(无真异步 I/O),确认 async 方法状态机本身是否分配。
        // 若 ≈ 同步 15KB → async Task 包装零分配,81KB 差距在异步 flush 链;若 ≈96KB → async 方法本身分配。
        [BenchmarkCategory("deflate-diag"), Benchmark]
        public async Task<long> Mio_100k_SyncWriteInAsync()
        {
            _stream100k.Reset();
            Xlsx.Write(_stream100k, _100k);
            await Task.CompletedTask;
            return _stream100k.Length;
        }

        // ---- 诊断:BCL DeflateStream sync vs async 的内部分配差(隔离异步层元凶) ----
        // 同样的可压缩输入 + MemoryStream(同步完成),测 DeflateStream.Write(span) vs WriteAsync(ROM)。
        // 若 async 远高于 sync,则异步路径固定开销主要落在 BCL DeflateStream 异步内部,非库可控。
        // 分块写(模拟 Mio 的 ~50 次 sheet flush),测 DeflateStream 每次 WriteAsync 的 per-call 分配。
        [BenchmarkCategory("deflate-diag"), Benchmark]
        public long Deflate_Sync()
        {
            _stream100k.Reset();
            using var ds = new DeflateStream(_stream100k, CompressionLevel.Optimal, leaveOpen: true);
            int chunk = _deflateInput.Length / 50;
            for (int i = 0; i < 50; i++)
            {
#if NETSTANDARD2_0
                ds.Write(_deflateInput, i * chunk, chunk);
#else
                ds.Write(_deflateInput.AsSpan(i * chunk, chunk));
#endif
            }
            ds.Dispose();
            return _stream100k.Length;
        }

        [BenchmarkCategory("deflate-diag"), Benchmark]
        public async Task<long> Deflate_Async()
        {
            _stream100k.Reset();
            using var ds = new DeflateStream(_stream100k, CompressionLevel.Optimal, leaveOpen: true);
            int chunk = _deflateInput.Length / 50;
            for (int i = 0; i < 50; i++)
            {
#if NETSTANDARD2_0
                await ds.WriteAsync(_deflateInput, i * chunk, chunk).ConfigureAwait(false);
#else
                await ds.WriteAsync(_deflateInput.AsMemory(i * chunk, chunk)).ConfigureAwait(false);
#endif
            }
            await ds.DisposeAsync().ConfigureAwait(false);
            return _stream100k.Length;
        }

        private sealed class NonSeekableWriteStream : Stream
        {
            private readonly Stream _inner;

            public NonSeekableWriteStream(Stream inner)
            {
                _inner = inner;
            }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new System.NotSupportedException();

            public override long Position
            {
                get => throw new System.NotSupportedException();
                set => throw new System.NotSupportedException();
            }

            public override void Flush() => _inner.Flush();

            public override int Read(byte[] buffer, int offset, int count) => throw new System.NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new System.NotSupportedException();

            public override void SetLength(long value) => _inner.SetLength(value);

            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

#if !NETSTANDARD2_0
            public override void Write(System.ReadOnlySpan<byte> buffer) => _inner.Write(buffer);
#endif
        }

        private sealed class FixedBufferWriter : IBufferWriter<byte>
        {
            private readonly byte[] _buffer;
            private int _written;

            public FixedBufferWriter(int capacity)
            {
                _buffer = new byte[capacity];
            }

            public int WrittenCount => _written;

            public void Reset() => _written = 0;

            public void Advance(int count)
            {
                var next = _written + count;
                if ((uint)next > (uint)_buffer.Length)
                    throw new InvalidOperationException("fixed buffer overflow");
                _written = next;
            }

            public Memory<byte> GetMemory(int sizeHint = 0)
            {
                EnsureCapacity(sizeHint);
                return _buffer.AsMemory(_written);
            }

            public Span<byte> GetSpan(int sizeHint = 0)
            {
                EnsureCapacity(sizeHint);
                return _buffer.AsSpan(_written);
            }

            private void EnsureCapacity(int sizeHint)
            {
                if (sizeHint < 0) throw new ArgumentOutOfRangeException(nameof(sizeHint));
                if (sizeHint == 0) sizeHint = 1;
                if (_written + sizeHint > _buffer.Length)
                    throw new InvalidOperationException("fixed buffer overflow");
            }
        }

        private sealed class FixedWriteStream : Stream
        {
            private readonly byte[] _buffer;
            private int _length;

            public FixedWriteStream(int capacity)
            {
                _buffer = new byte[capacity];
            }

            public override long Length => _length;

            public void Reset() => _length = 0;

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Position
            {
                get => _length;
                set => throw new NotSupportedException();
            }

            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count)
            {
                if ((uint)offset > (uint)buffer.Length) throw new ArgumentOutOfRangeException(nameof(offset));
                if ((uint)count > (uint)(buffer.Length - offset)) throw new ArgumentOutOfRangeException(nameof(count));
                EnsureCapacity(count);
                Buffer.BlockCopy(buffer, offset, _buffer, _length, count);
                _length += count;
            }

#if !NETSTANDARD2_0
            public override void Write(ReadOnlySpan<byte> buffer)
            {
                EnsureCapacity(buffer.Length);
                buffer.CopyTo(_buffer.AsSpan(_length));
                _length += buffer.Length;
            }

            // 关键:override WriteAsync(ROM) 高效版,避免 Stream 基类默认实现的 ValueTask 包装分配。
            // 基类默认 WriteAsync(ROM) 调 WriteAsync(byte[])+Task 包装,每次分配;override 后同步完成零分配。
            public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                Write(buffer.Span);
                return default;
            }
#endif

            private void EnsureCapacity(int count)
            {
                if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
                if (_length + count > _buffer.Length)
                    throw new InvalidOperationException("fixed stream overflow");
            }
        }

        // 真异步 I/O mock:每次 WriteAsync 先 Task.Yield 再写,模拟 NetworkStream/FileStream 异步完成。
        // 用于验证零分配改造(AsTask 消除)、批量化 await、IList 快路径在异步完成场景的实际收益
        // (FixedWriteStream 同步完成下 AsTask 返回缓存 Task 不分配,测不出这些优化的真实效果)。
        private sealed class AsyncYieldingWriteStream : Stream
        {
            private readonly Stream _inner;
            public AsyncYieldingWriteStream(Stream inner) => _inner = inner;

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

            public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                await Task.Yield();
                _inner.Write(buffer, offset, count);
            }

#if !NETSTANDARD2_0
            public override void Write(ReadOnlySpan<byte> buffer) => _inner.Write(buffer);

            public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                await Task.Yield();
                _inner.Write(buffer.Span);
            }
#endif
        }
    }
}
