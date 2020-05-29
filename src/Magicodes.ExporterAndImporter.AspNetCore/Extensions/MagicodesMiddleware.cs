using Magicodes.ExporterAndImporter.Attributes;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Extensions
{
    public class MagicodesMiddleware
    {
        private readonly RequestDelegate _next;
        private readonly ILogger<MagicodesMiddleware> _logger;
        private readonly MagicodesBase _extensions;
        public MagicodesMiddleware(RequestDelegate next, ILogger<MagicodesMiddleware> logger)
        {
            _next = next;
            _logger = logger;
            _extensions = new MagicodesBase();
        }
        public async Task InvokeAsync(HttpContext context)
        {
            using var memoryStream = new MemoryStream();
            var originalResponseBodyStream = context.Response.Body;
            try
            {
                var endpoint = context.GetEndpoint();
                var endpointMagicodesData = endpoint?.Metadata.GetMetadata<IMagicodesData>();
                if (endpointMagicodesData != null)
                {
                    context.Response.Body = memoryStream;
                    await _next.Invoke(context);
                    context.Response.Body = originalResponseBodyStream;
                    var bodyAsText = await _extensions.ReadResponseBodyStreamAsync(memoryStream);
                    var isSuccessful = await _extensions.HandleSuccessfulReqeustAsync(context: context, body: bodyAsText, tplPath: endpointMagicodesData.TemplatePath,
                        type: endpointMagicodesData.Type);
                    if (!isSuccessful)
                    {
                        memoryStream.Seek(0, SeekOrigin.Begin);
                        await memoryStream.CopyToAsync(originalResponseBodyStream);
                    }
                }
                else
                {
                    await _next(context);
                }
            }
            catch
            {
                memoryStream.Seek(0, SeekOrigin.Begin);
                await memoryStream.CopyToAsync(originalResponseBodyStream);
            }
            finally
            {
                _logger.Log(LogLevel.Information, $@"Source:[{context.Connection.RemoteIpAddress }] 
                                                     Request: {context.Request.Method} {context.Request.Scheme} {context.Request.Host}{context.Request.Path} {context.Request.QueryString}
                                                     Responded with [{context.Response.StatusCode}] ");
            }

        }
    }
}
