using Microsoft.Extensions.DependencyInjection;

namespace Magicodes.ExporterAndImporter.Pdf
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddMagicodesPdfExporter(this IServiceCollection services)
        {
            services.AddSingleton<IPdfNativeLibraryService, PdfNativeLibraryService>();
            services.AddScoped<IPdfExporter, PdfExporter>();
            return services;
        }
    }
}
