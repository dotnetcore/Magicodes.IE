using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Excel;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc.Formatters;
using Microsoft.Net.Http.Headers;

namespace Magicodes.ExporterAndImporter.ASPCore.Formatters
{
    public class ExcelOuputFormatter : OutputFormatter
    {
        protected ExcelOuputFormatter()
        {
            SupportedMediaTypes.Clear();
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.ms-excel"));
        }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override Task WriteResponseBodyAsync(OutputFormatterWriteContext context)
        {
            //context.HttpContext.Response.
        }

        public override void WriteResponseHeaders(OutputFormatterWriteContext context)
        {
            // Get the raw request URI.
            string rawUri = context.HttpContext.Request.GetDisplayUrl();


            // Remove query string if present.
            int queryStringIndex = rawUri.IndexOf('?');
            if (queryStringIndex > -1)
            {
                rawUri = rawUri.Substring(0, queryStringIndex);
            }

            string fileName;

            
            // Look for ExcelDocumentAttribute on class.
            var itemType = context.ObjectType.GetEnumerableItemType();
            var excelDocumentAttribute = (itemType ?? context.ObjectType).GetAttribute<ExcelExporterAttribute>();

            if (excelDocumentAttribute != null && !string.IsNullOrEmpty(excelDocumentAttribute.Name))
            {
                fileName = excelDocumentAttribute.Name;
            }
            else
            {
                fileName = Path.GetFileName(rawUri) ?? "data";
                if (fileName.Contains("?")) fileName = fileName.Split('?')[0];
            }

            // 添加 XLSX 扩展名
            if (!fileName.EndsWith("xlsx", StringComparison.CurrentCultureIgnoreCase)) fileName += ".xlsx";

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = fileName,
                Inline = true  // false = prompt the user for downloading;  true = browser to try to show the file inline
            };
            context.HttpContext.Response.Headers.Add("Content-Disposition", cd.ToString());
        }
    }
}
