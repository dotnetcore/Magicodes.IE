using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Diagnostics.Windows.Configs;
using BenchmarkDotNet.Jobs;
using Magicodes.Benchmarks.Models;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Pdf;

namespace Magicodes.Benchmarks
{
    //[TailCallDiagnoser]
    //[EtwProfiler]
    [SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5, runtimeMoniker: RuntimeMoniker.NetCoreApp31)]
    [MemoryDiagnoser]
    public class ExportPdfBenchmarks
    {
        [Params(100, 1000, 10000)]
        public int RowsCount;
        private readonly List<ExportTestData> _exportTestData = new List<ExportTestData>();
        private readonly PdfExporter _exporter;
        private string tpl;

        public ExportPdfBenchmarks()
        {
            _exporter = new PdfExporter();
        }

        [GlobalSetup]
        public void GlobalSetup()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "tpl1.cshtml");
            tpl = File.ReadAllText(tplPath);
            for (var i = 1; i <= RowsCount; i++)
            {
                _exportTestData.Add(new ExportTestData
                {
                    Name1 = "1",
                    Name2 = "2",
                    Name3 = "3",
                    Name4 = "4"
                });
            }
        }


        [GlobalCleanup]
        public void GlobalCleanup()
        {
            _exportTestData.Clear();
            tpl = null;
        }

        [Benchmark]
        public async Task ExportPdfAsByteArrayTest()
        {
            await _exporter.ExportListBytesByTemplate(_exportTestData, tpl);
        }

    }
}
