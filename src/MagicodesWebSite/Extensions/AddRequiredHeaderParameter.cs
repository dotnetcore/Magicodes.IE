using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;
using System.Collections.Generic;

namespace MagicodesWebSite.Extensions
{
    public class AddRequiredHeaderParameter : IOperationFilter
    {
        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {
            if (operation.Parameters == null)
            {
                operation.Parameters = new List<OpenApiParameter>();
            }

            operation.Parameters.Add(new OpenApiParameter
            {
                Name = "Magicodes-Type",
                In = ParameterLocation.Header,
                Required = false,
                Description = "根据HttpContentMediaType添加指定的header值，导出不同格式的文件。"
            });
        }
    }
}
