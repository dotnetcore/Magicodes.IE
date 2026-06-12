using System;
using System.Runtime.InteropServices;
using Magicodes.ExporterAndImporter.Pdf;
using Shouldly;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class PdfNativeLibraryBootstrapper_Tests
    {
        [Fact]
        public void CheckEnvironment_ReturnsValidDiagnostics()
        {
            var info = PdfEnvironmentInfo.Check();

            info.Platform.ShouldNotBeNullOrWhiteSpace();
            info.InstallSuggestion.ShouldNotBeNullOrWhiteSpace();

            if (info.IsAvailable)
            {
                info.NativeLibraryPath.ShouldNotBeNullOrWhiteSpace();
                info.NativeLibraryFileExists.ShouldBeTrue();
                info.NativeLibraryLoadable.ShouldBeTrue();
                info.ErrorDetail.ShouldBeNull();
            }
            else
            {
                info.ErrorDetail.ShouldNotBeNullOrWhiteSpace();
            }
        }

        [Fact]
        public void CheckEnvironment_ToString_ContainsPlatformAndDiagnostics()
        {
            var info = PdfEnvironmentInfo.Check();
            var output = info.ToString();

            output.ShouldNotBeNullOrWhiteSpace();
            output.ShouldContain(info.Platform);

            if (!info.IsAvailable)
            {
                output.ShouldContain("[FAIL]");
                output.ShouldContain("Fix:");
            }
        }

        [Fact]
        public void CheckEnvironment_Platform_MatchesRuntimeIdentifier()
        {
            var info = PdfEnvironmentInfo.Check();
#if NET5_0_OR_GREATER
            info.Platform.ShouldBe(RuntimeInformation.RuntimeIdentifier);
#else
            info.Platform.ShouldNotBeNullOrWhiteSpace();
#endif
        }
    }
}
