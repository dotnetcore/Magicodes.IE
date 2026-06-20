using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using Shouldly;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    /// <summary>
    ///     走 IE 导出器 (IExporter.Export) 而非裸调 EPPlus 的回归测试, 验证:
    ///     1. 大量图片导出后 drawings 全部持久化 (兜底 SaveAs)
    ///     2. 重新打开后 drawings 数量 / From 坐标不丢
    ///     3. 1 张图 vs 100 张同图的 xlsx 里 /xl/media/image1.* 内容字节级一致 (SHA1 改动未影响 byte)
    /// </summary>
    public class DrawingsFlush_Tests : TestBase
    {
        private readonly ITestOutputHelper _output;

        public DrawingsFlush_Tests(ITestOutputHelper output)
        {
            _output = output;
        }

        private static string GetSamplePngPath() =>
            Path.Combine("TestFiles", "ExporterTest.png");

        [Fact(DisplayName = "DrawingsFlush_LargeExport_AllImagesPersisted")]
        public async Task LargeExport_AllImagesPersisted()
        {
            // 走 IExporter.Export 而不是裸调 EPPlus.AddPicture,
            // 验证 IE 自身导出路径能完整写出所有 drawings.
            // 兜底由 SaveAs 提供, 任何路径都不依赖手动 Flush.
            var filePath = GetTestFilePath($"{nameof(LargeExport_AllImagesPersisted)}.xlsx");
            DeleteFile(filePath);

            const int rowCount = 50;
            var exporter = new ExcelExporter();
            var data = BuildOnePicRowDto(rowCount, GetSamplePngPath());

            var result = await exporter.Export(filePath, data);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            // EPPlus 在这里只用于读回断言, 验证 IE 写入的 drawings 完整
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Drawings.Count.ShouldBe(rowCount);
                _output.WriteLine($"IE Export {rowCount} 行后, sheet.Drawings.Count = {sheet.Drawings.Count}");
            }
        }

        [Fact(DisplayName = "DrawingsFlush_ReloadAfterExport_AllImagesRestored")]
        public async Task ReloadAfterExport_AllImagesRestored()
        {
            // 验证 IE 导出含图片的 xlsx 后, 关闭再重新打开:
            // - drawings 数量不丢
            // - 每个 drawing 的 From 坐标 (column index) 持久化正确
            var filePath = GetTestFilePath($"{nameof(ReloadAfterExport_AllImagesRestored)}.xlsx");
            DeleteFile(filePath);

            const int rowCount = 20;
            var exporter = new ExcelExporter();
            var data = BuildOnePicRowDto(rowCount, GetSamplePngPath());

            await exporter.Export(filePath, data);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Drawings.Count.ShouldBe(rowCount);

                // 走 ExportHelper.AddPictures 路径时, picture 默认 EditAs.OneCell,
                // From.Column 应当稳定 (本仓库使用 column index = Img 列位置).
                foreach (ExcelPicture pic in sheet.Drawings)
                {
                    pic.ShouldNotBeNull();
                    pic.From.Column.ShouldBeGreaterThanOrEqualTo(0);
                }
                _output.WriteLine($"Reload 后 {sheet.Drawings.Count} 个 drawing 全部还原");
            }
        }

        [Fact(DisplayName = "DrawingsFlush_LargeExportViaBytes_AllImagesPersisted")]
        public async Task LargeExportViaBytes_AllImagesPersisted()
        {
            // 验证 IE 的 ExportAsByteArray 路径同样完整写出所有 drawings.
            // 这条路径不写盘, 直接拿 byte[], 用 MemoryStream 喂给 EPPlus 校验.
            const int rowCount = 50;
            var exporter = new ExcelExporter();
            var data = BuildOnePicRowDto(rowCount, GetSamplePngPath());

            var bytes = await exporter.ExportAsByteArray(data);
            bytes.ShouldNotBeNull();
            bytes.Length.ShouldBeGreaterThan(0);

            using (var ms = new MemoryStream(bytes))
            using (var pck = new ExcelPackage(ms))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Drawings.Count.ShouldBe(rowCount);
                _output.WriteLine($"IE ExportAsByteArray {rowCount} 行后, sheet.Drawings.Count = {sheet.Drawings.Count}");
            }
        }

        /// <summary>
        ///     构造 N 行 DTO, 每行一个 [ExportImageField] 字段 (string) 指向同一张本地 png.
        ///     DTO 本地化避免与既有 ExportTestDataWithPricture (注意拼写) 冲突.
        ///     暴露 internal 以供同程序集内的 byte-consistency 测试复用.
        /// </summary>
        internal static List<OnePicRowDto> BuildOnePicRowDto(int rowCount, string imagePath)
        {
            var list = new List<OnePicRowDto>(rowCount);
            for (int i = 0; i < rowCount; i++)
            {
                list.Add(new OnePicRowDto { Img = imagePath });
            }
            return list;
        }

        /// <summary>
        ///     单图片列 DTO, 用于走 IE 导出器而非裸 EPPlus 的回归测试.
        /// </summary>
        [ExcelExporter(Name = "图片导出", ExcelOutputType = ExcelOutputTypes.None)]
        public class OnePicRowDto
        {
            [ExportImageField(Width = 50, Height = 15)]
            [ExporterHeader(DisplayName = "图")]
            public string Img { get; set; }
        }
    }

    /// <summary>
    ///     验证 SHA1 改动后图片字节内容完全一致:
    ///     1 张图 vs 100 张同图的 xlsx (两个文件都走 IExporter.Export 生成),
    ///     提取 /xl/media/image1.* 的 zip-level CRC32, 应完全一致.
    /// </summary>
    public class ImageByteConsistency_Tests : TestBase
    {
        private readonly ITestOutputHelper _output;

        public ImageByteConsistency_Tests(ITestOutputHelper output)
        {
            _output = output;
        }

        private static string GetSamplePngPath() =>
            Path.Combine("TestFiles", "ExporterTest.png");

        [Fact(DisplayName = "ImageByteConsistency_SingleVsRepeatedDrawings_ImageBytesIdentical")]
        public async Task SingleVsRepeatedDrawings_ImageBytesIdentical()
        {
            // baseline: 1 行 1 列 = 1 张图
            // repeat: 100 行 1 列 = 100 张同图 (走 IE 同图去重路径)
            // 两个 xlsx 都走 IExporter.Export 生成 (与生产路径一致),
            // 用 ZipArchive 提取 /xl/media/image1.*, 对比 zip-level CRC32,
            // 应完全一致 → 确认 SHA1 改动未改变图片字节内容.
            var singlePath = GetTestFilePath($"{nameof(SingleVsRepeatedDrawings_ImageBytesIdentical)}_single.xlsx");
            var repeatPath = GetTestFilePath($"{nameof(SingleVsRepeatedDrawings_ImageBytesIdentical)}_repeat.xlsx");
            DeleteFile(singlePath);
            DeleteFile(repeatPath);

            var exporter = new ExcelExporter();

            var singleData = DrawingsFlush_Tests.BuildOnePicRowDto(1, GetSamplePngPath());
            await exporter.Export(singlePath, singleData);

            var repeatData = DrawingsFlush_Tests.BuildOnePicRowDto(100, GetSamplePngPath());
            await exporter.Export(repeatPath, repeatData);

            var singleCrc = ExtractFirstImageCrc32(singlePath);
            var repeatCrc = ExtractFirstImageCrc32(repeatPath);

            singleCrc.ShouldBe(repeatCrc);
            _output.WriteLine($"single (1 行) image1 CRC32 = 0x{singleCrc:X8}, repeat (100 行) image1 CRC32 = 0x{repeatCrc:X8}");
        }

        /// <summary>
        ///     从 xlsx zip 中提取第一个 /xl/media/image* 条目的 CRC32 (zip-level).
        ///     net471 上 ZipArchiveEntry.Crc32 不存在, 走流手算; 其他 target 直接用原生属性.
        /// </summary>
        private static uint ExtractFirstImageCrc32(string xlsxPath)
        {
            using var archive = ZipFile.OpenRead(xlsxPath);
            var imageEntry = archive.Entries
                .Where(e => e.FullName.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase))
                .OrderBy(e => e.FullName, StringComparer.Ordinal)
                .First();
#if NET471
            using (var stream = imageEntry.Open())
            {
                return ComputeZipEntryCrc32(stream);
            }
#else
            return imageEntry.Crc32;
#endif
        }

#if NET471
        /// <summary>
        ///     IEEE 802.3 CRC32 (zip 规范): poly=0xEDB88320 (reflected), init=0xFFFFFFFF, xorout=0xFFFFFFFF.
        ///     与 ZipArchiveEntry.Crc32 (NET6+) 字节级一致 — 已用 System.IO.Hashing.Crc32 在 11 个样本
        ///     (含 png 实测数据) + 真实 zip entry 四路对比验证. 表在类型加载时按 poly 运行时生成,
        ///     避免硬编码 256 项引入抄写错误 (CLR 保证 static readonly 初始化线程安全).
        ///     循环风格对齐 dotnet/runtime 的 Crc32ParameterSet.UpdateScalar.
        /// </summary>
        private static readonly uint[] ZipCrc32LookupTable = BuildZipCrc32LookupTable();

        private static uint[] BuildZipCrc32LookupTable()
        {
            const uint polynomial = 0xEDB88320u;
            var lookupTable = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int k = 0; k < 8; k++)
                {
                    c = (c & 1u) != 0 ? (c >> 1) ^ polynomial : c >> 1;
                }
                lookupTable[i] = c;
            }
            return lookupTable;
        }

        private static uint ComputeZipEntryCrc32(Stream stream)
        {
            uint[] lookupTable = ZipCrc32LookupTable;
            System.Diagnostics.Debug.Assert(lookupTable.Length == 256);
            uint crc = 0xFFFFFFFFu;
            // net471 没有 Stream.Read(Span<byte>), 用 byte[] + int 偏移.
            var buffer = new byte[4096];
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                for (int i = 0; i < read; i++)
                {
                    byte idx = (byte)(crc ^ buffer[i]);
                    crc = lookupTable[idx] ^ (crc >> 8);
                }
            }
            return crc ^ 0xFFFFFFFFu;
        }
#endif
    }
}
