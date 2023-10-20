using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using Magicodes.Benchmarks.Models;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.Benchmarks
{
    [SimpleJob(launchCount: 2, warmupCount: 2, targetCount: 5, runtimeMoniker: RuntimeMoniker.NetCoreApp50)]
    [MemoryDiagnoser]
    [ThreadingDiagnoser]
    //[SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5, runtimeMoniker: RuntimeMoniker.NetCoreApp22)]
    //[SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 5, runtimeMoniker: RuntimeMoniker.NetCoreApp31)]
    public class ImportExcelBenchmarks
    {
        private readonly List<ImportStudentDto> _studentDtos = new List<ImportStudentDto>();

        [Params(240000,10,1000,10000,500000,1000000)]
        public int RowsCount;

        private Stream _stream;
        private readonly IImporter _importer;

        public ImportExcelBenchmarks()
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
                    StudentCode = "A",
                    A = "1",
                    A1 = "2",
                    A2 = "3",
                    A3 = "4",
                });
            }
            IExporter exporter = new ExcelExporter();
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
            // await _importer.Import<ImportStudentDto>(_stream);
            using (var importer = new Magicodes.ExporterAndImporter.Excel.Utility.ImportHelper<ImportStudentDto>(_stream, null))
            {
                var data = await importer.Import();
            }
        }

        [Benchmark]
        public async Task ImportByStreamTest1()
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
