using Magicodes.ExporterAndImporter.Pdf;
using Microsoft.Extensions.DependencyInjection;
using Shouldly;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class PdfServiceCollectionExtensions_Tests
    {
        [Fact]
        public void AddMagicodesPdfExporter_RegistersExpectedServices()
        {
            var services = new ServiceCollection();
            services.AddMagicodesPdfExporter();

            var provider = services.BuildServiceProvider();

            provider.GetService<IPdfExporter>().ShouldNotBeNull();
            provider.GetService<IPdfNativeLibraryService>().ShouldNotBeNull();
        }
    }
}
