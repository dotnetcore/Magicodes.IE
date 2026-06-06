using Magicodes.ExporterAndImporter.Attributes;
using Magicodes.ExporterAndImporter.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Filters
{
    public class MagicodesFilter : IAsyncResultFilter
    {
        public async Task OnResultExecutionAsync(ResultExecutingContext context, ResultExecutionDelegate next)
        {
            var result = context.Result;
            if (result is ObjectResult objectResult)
            {
                var endpoint = context.HttpContext.GetEndpoint();
                var endpointMagicodesData = endpoint?.Metadata.GetMetadata<IMagicodesData>();
                if (endpointMagicodesData != null)
                {
                    var extensions = context.HttpContext.RequestServices.GetService<MagicodesBase>()
                        ?? new MagicodesBase(context.HttpContext.RequestServices);
                    var timeConverter = new IsoDateTimeConverter { DateTimeFormat = "yyyy-MM-dd HH:mm:ss" };
                    var json = JsonConvert.SerializeObject(objectResult.Value, timeConverter);
                    var isSuccessful = await extensions.HandleSuccessfulReqeustAsync(context: context.HttpContext, body: json, tplPath: endpointMagicodesData.TemplatePath,
                        type: endpointMagicodesData.Type);
                    if (!isSuccessful)
                    {
                        await next();
                    }
                }
            }
            else
            {
                await next();
            }
        }

    }
}
