#if NETCOREAPP3_0 || NETCOREAPP3_1 || NET5_0 || NET6_0_OR_GREATER || NET7_0_OR_GREATER || NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Extensions;
using Magicodes.ExporterAndImporter.Pdf;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Shouldly;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class MagicodesBase_Tests
    {
        private const string PdfContentType = "application/pdf";

        [Fact]
        public async Task HandleSuccessfulReqeustAsync_UsesRegisteredPdfExporter_WhenAvailable()
        {
            var services = new ServiceCollection();
            var exporter = new StubPdfExporter();
            services.AddSingleton<IPdfExporter>(exporter);

            var context = new DefaultHttpContext
            {
                RequestServices = services.BuildServiceProvider()
            };
            context.Request.Headers["Magicodes-Type"] = PdfContentType;
            context.Response.Body = new MemoryStream();

            var templatePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".cshtml");
            File.WriteAllText(templatePath, "<p>Hello</p>");

            try
            {
                var extensions = new MagicodesBase();
                var handled = await extensions.HandleSuccessfulReqeustAsync(context, "{\"Name\":\"Magicodes\"}", typeof(StubPdfModel), templatePath);

                handled.ShouldBeTrue();
                context.Response.ContentType.ShouldBe(PdfContentType);
                exporter.ExportByObjectCallCount.ShouldBe(1);
                ((MemoryStream)context.Response.Body).ToArray().ShouldBe(exporter.ExpectedBytes);
            }
            finally
            {
                if (File.Exists(templatePath))
                {
                    File.Delete(templatePath);
                }
            }
        }

        private class StubPdfModel
        {
            public string Name { get; set; }
        }

        private class StubPdfExporter : IPdfExporter
        {
            public byte[] ExpectedBytes { get; } = { 1, 2, 3, 4 };

            public int ExportByObjectCallCount { get; private set; }

            public Task<ExportFileInfo> ExportListByTemplate<T>(string fileName, ICollection<T> dataItems, string htmlTemplate = null) where T : class
            {
                throw new NotImplementedException();
            }

            public Task<ExportFileInfo> ExportByTemplate<T>(string fileName, T data, string template) where T : class
            {
                throw new NotImplementedException();
            }

            public Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class
            {
                throw new NotImplementedException();
            }

            public Task<byte[]> ExportBytesByTemplate(object data, string template, Type type)
            {
                ExportByObjectCallCount++;
                return Task.FromResult(ExpectedBytes);
            }

            public Task<byte[]> ExportListBytesByTemplate<T>(ICollection<T> data, PdfExportOptions pdfExportOptions, string template) where T : class
            {
                throw new NotImplementedException();
            }

            public Task<byte[]> ExportBytesByTemplate<T>(T data, PdfExportOptions pdfExportOptions, string template) where T : class
            {
                throw new NotImplementedException();
            }

            public Task<byte[]> ExportListBytesByTemplate<T>(ICollection<T> data, PdfExporterAttribute pdfExporterAttribute, string template) where T : class
            {
                throw new NotImplementedException();
            }

            public Task<byte[]> ExportBytesByTemplate<T>(T data, PdfExporterAttribute pdfExporterAttribute, string template) where T : class
            {
                throw new NotImplementedException();
            }
        }
    }
}
#endif
