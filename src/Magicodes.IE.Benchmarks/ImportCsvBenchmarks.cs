using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using Magicodes.Benchmarks.Models;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Csv;
using Magicodes.ExporterAndImporter.Excel;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.Benchmarks
{
    [SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5, runtimeMoniker: RuntimeMoniker.Net461)]
    [SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5, runtimeMoniker: RuntimeMoniker.NetCoreApp22)]
    [SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5, runtimeMoniker: RuntimeMoniker.NetCoreApp31)]
    public class ImportCsvBenchmarks
    {
        private readonly List<ImportStudentDto> _studentDtos = new List<ImportStudentDto>();

        [Params(10000, 120000, 240000, 500000, 1000000)]
        public int RowsCount;

        private Stream _stream;
        private readonly IImporter _importer;

        public ImportCsvBenchmarks()
        {
            _importer = new ExcelImporter();
        }

        [GlobalSetup]
        public async Task GlobalSetup()
        {
            for (var i = 1; i <= RowsCount; i++)
            {
                _studentDtos.Add(new ImportStudentDto
                {
                    Phone = "13211111111",
                    Name = "Mr.A",
                    IdCard = "1111111111111111111",
                    SerialNumber = i,
                    StudentCode = "A"
                });
            }
            IExporter exporter = new CsvExporter();
            _stream = BytesToStream(await exporter.ExportAsByteArray(_studentDtos));
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            _studentDtos.Clear();
            _stream = null;
        }

        [Benchmark]
        public async Task ImportByStreamTest()
        {
            await _importer.Import<ImportStudentDto>(_stream);
        }

        /// <summary> 
        /// 将 byte[] 转成 Stream 
        /// </summary> 
        public Stream BytesToStream(byte[] bytes)
        {
            Stream stream = new MemoryStream(bytes);
            return stream;
        }
    }
}
