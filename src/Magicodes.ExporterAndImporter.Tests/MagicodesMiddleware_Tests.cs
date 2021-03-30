#if NETCOREAPP3_0 || NETCOREAPP3_1 || NET5_0
using Magicodes.ExporterAndImporter.Builder;
using MagicodesWebSite;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.TestHost;
using OfficeOpenXml;
using Shouldly;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class MagicodesMiddleware_Tests
    {
        private readonly HttpClient _client;
        private readonly TestServer _testServer;

        public MagicodesMiddleware_Tests()
        {
            var webHostBuilder = Program.CreateHostBuilder(new[] { Program.MiddlewareScenario });
            _testServer = new TestServer(webHostBuilder);
            _client = _testServer.CreateClient();
            _client.BaseAddress = new Uri("http://localhost");
        }
        [Fact(DisplayName = "空文件类型头")]
        public async Task AllowsEmptyFileHttpContentMediaType()
        {
            var text = "Hello world";
            var builder = new WebHostBuilder()
                .Configure(app =>
                {
                    app.UseMagiCodesIE();
                    app.Run(context => context.Response.WriteAsync("Hello world"));
                });
            var server = new TestServer(builder);
            var client = server.CreateClient();
            client.BaseAddress = new Uri("https://example.com:5050");

            var request = new HttpRequestMessage(HttpMethod.Get, "");

            var response = await client.SendAsync(request);
            var repstr = await response.Content.ReadAsStringAsync();
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.Equal(text, repstr);
        }

#region FunctionalTests
        [Fact]
        public async Task AllowsXlsxHttpContentMediaType()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/excel");
            var expectedContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);
        }

        [Fact]
        public async Task AllowsPDFHttpContentMediaType()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/pdf");
            var expectedContentType = "application/pdf";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);
        }
        [Fact]
        public async Task AllowsDocxHttpContentMediaType()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/Word");
            var expectedContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);
        }
        [Fact]
        public async Task AllowsHtmlHttpContentMediaType()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/html");
            var expectedContentType = "text/html";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);
        }

        [Fact]
        public async Task XlsxHttpContentMediaType_AttrsExport_Test()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/excel");
            var expectedContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);

            var result = await response.Content.ReadAsByteArrayAsync();
            result.ShouldNotBeNull();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(XlsxHttpContentMediaType_AttrsExport_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);
            using (var file = File.OpenWrite(filePath))
            {
                file.Write(result, 0, result.Length);
            }
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells[sheet.Dimension.Address].Rows.ShouldBe(101);

                //[ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
                sheet.Cells["E2"].Text.Equals(DateTime.Parse(sheet.Cells["E2"].Text).ToString("yyyy-MM-dd"));

                //[ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
                sheet.Cells["F2"].Text.Equals(DateTime.Parse(sheet.Cells["F2"].Text).ToString("yyyy-MM-dd HH:mm:ss"));

                //默认DateTime
                sheet.Cells["G2"].Text.Equals(DateTime.Parse(sheet.Cells["G2"].Text).ToString("yyyy-MM-dd"));

                sheet.Tables.Count.ShouldBe(1);

                var tb = sheet.Tables.First();
                tb.Columns.Count.ShouldBe(9);
                tb.Columns.First().Name.ShouldBe("加粗文本");
            }
        }
        [Fact]
        public async Task PdfHttpContentMediaType_BathExportPortraitReceipt_Test()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/pdf");
            var expectedContentType = "application/pdf";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);
            var result = await response.Content.ReadAsByteArrayAsync();
            result.ShouldNotBeNull();
        }
        [Fact]
        public async Task DocxHttpContentMediaType_ExportWordFileByTemplate_Test()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/Word");
            var expectedContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);
            var result = await response.Content.ReadAsByteArrayAsync();
            result.ShouldNotBeNull();
        }
        [Fact]
        public async Task HtmlHttpContentMediaType_ExportReceipt_Test()
        {
            // Arrange
            var message = new HttpRequestMessage(HttpMethod.Get, $"/api/Magicodes/html");
            var expectedContentType = "text/html";
            message.Headers.Add("Magicodes-Type", expectedContentType);
            // Act
            var response = await _client.SendAsync(message);
            // Assert
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            Assert.NotNull(response.Content);
            Assert.NotNull(response.Content.Headers.ContentType);
            Assert.Equal(expectedContentType, response.Content.Headers.ContentType.MediaType);
            var result = await response.Content.ReadAsByteArrayAsync();
            result.ShouldNotBeNull();
        }

#endregion


    }
}
#endif
