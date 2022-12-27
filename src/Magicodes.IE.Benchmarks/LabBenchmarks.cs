using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Diagnostics.Windows.Configs;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Jobs;
using Magicodes.Benchmarks.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Excel.Utility;
using MiniExcelLibs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.Benchmarks
{
    [ThreadingDiagnoser]
    [TailCallDiagnoser]
    [MemoryDiagnoser]
    [SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5, runtimeMoniker: RuntimeMoniker.NetCoreApp50)]
    public class LabBenchmarks
    {
        [Params(100, 500, 1000, 2000, 3000, 5000, 10000, 100000)]
        public int RowsCount;
        private readonly List<ExportTestDataWithAttrs> _exportTestData = new List<ExportTestDataWithAttrs>();
        private readonly IExporter _exporter;

        public LabBenchmarks()
        {
        }

        [GlobalSetup]
        public void GlobalSetup()
        {
            for (var i = 1; i <= RowsCount; i++)
            {
                _exportTestData.Add(new ExportTestDataWithAttrs
                {
                    Age = i,
                    Name = "Mr.A",
                    Text3 = "Text3"
                });
            }
        }


        [GlobalCleanup]
        public void GlobalCleanup()
        {
            _exportTestData.Clear();
        }

        [Benchmark]
        public async Task ExportExcelAsByteArrayAsyncTest()
        {
            var helper = new ExportHelper<ExportTestDataWithAttrs>();

            using (helper.CurrentExcelPackage)
            {
                //helper.AddExcelWorksheet();
                helper.Export(_exportTestData);

                await helper.CurrentExcelPackage.GetAsByteArrayAsync();
            }
        }

        [Benchmark]
        public async Task ExportMiniExcelAsyncTest()
        {
            var memoryStream = new MemoryStream();
            await memoryStream.SaveAsAsync(_exportTestData);
            memoryStream.Seek(0, SeekOrigin.Begin);
        }
    }
}
