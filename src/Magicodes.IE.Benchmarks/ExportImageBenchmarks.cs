using BenchmarkDotNet.Attributes;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.Benchmarks
{
    [MemoryDiagnoser]
    [SimpleJob(launchCount: 1, warmupCount: 2, invocationCount: 5)]
    public class ExportImageBenchmarks
    {
        // Minimal valid 1x1 PNG (red pixel). Generated once in static ctor
        // and written to a temp file; EPPlus's AddPicture needs a FileInfo.
        private static readonly string SamplePngPath =
            Path.Combine(Path.GetTempPath(), "magicodes_benchmark_sample.png");

        [Params(100, 500)]
        public int RowCount;

        private List<ExportImageDto> _data;
        private IExcelExporter _exporter;

        static ExportImageBenchmarks()
        {
            File.WriteAllBytes(SamplePngPath, new byte[]
            {
                0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
                0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
                0x54, 0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
                0x00, 0x03, 0x01, 0x01, 0x00, 0xC9, 0xFE, 0x92,
                0xEF, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
                0xAE, 0x42, 0x60, 0x82
            });
        }

        [GlobalSetup]
        public void Setup()
        {
            _data = new List<ExportImageDto>(RowCount);
            for (var i = 0; i < RowCount; i++)
            {
                _data.Add(new ExportImageDto
                {
                    Name = $"Item {i}",
                    Image = SamplePngPath,
                });
            }
            _exporter = new ExcelExporter();
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            _data?.Clear();
        }

        [Benchmark(Description = "图片导出-ByteArray输出")]
        public async Task<byte[]> ExportImageAsByteArray()
        {
            return await _exporter.ExportAsByteArray(_data);
        }

        [Benchmark(Description = "图片导出-文件输出")]
        public async Task<string> ExportImageToFile()
        {
            var path = Path.Combine(Path.GetTempPath(), $"benchmark_img_{Guid.NewGuid()}.xlsx");
            await _exporter.Export(path, _data);
            if (File.Exists(path)) File.Delete(path);
            return path;
        }

        [Benchmark(Description = "图片导出-仅EPPlus(无IE包装)")]
        public async Task<byte[]> ExportEpplusOnly()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Images");
            for (var i = 0; i < RowCount; i++)
            {
                ws.Cells[i + 1, 1].Value = _data[i].Name;
                var pic = ws.Drawings.AddPicture($"Pic{i}", new FileInfo(SamplePngPath));
                pic.SetPosition(i, 0, 1, 0);
            }
            return await package.GetAsByteArrayAsync();
        }
    }

    public class ExportImageDto
    {
        public string Name { get; set; }
        [ExportImageField(Width = 50, Height = 50)]
        public string Image { get; set; }
    }
}
