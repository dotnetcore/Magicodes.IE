#if NETCOREAPP3_0||NETCOREAPP3_1
using Magicodes.ExporterAndImporter.Builder;
using MagicodesWebSite;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.TestHost;
using System;
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
            Assert.Equal(text,repstr);
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
        #endregion


    }
}
#endif
