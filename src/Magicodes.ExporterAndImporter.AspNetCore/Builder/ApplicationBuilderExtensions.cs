using Magicodes.ExporterAndImporter.Extensions;
using Microsoft.AspNetCore.Builder;
using System;

namespace Magicodes.ExporterAndImporter.Builder
{
    public static class ApplicationBuilderExtensions
    {
        /// <summary>
        /// Adds the <see cref="MagicodesMiddleware"/> to automatically set the export file
        /// requests based on information provided by the client.
        /// </summary>
        /// <param name="app">The <see cref="IApplicationBuilder"/>.</param>
        /// <returns>The <see cref="IApplicationBuilder"/>.</returns>
        public static IApplicationBuilder UseMagiCodesIE(this IApplicationBuilder app)
        {
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app));
            }

            return app.UseMiddleware<MagicodesMiddleware>();
        }
    }
}
