using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Magicodes.IE.IO;
using Xunit;

namespace Magicodes.IE.IO.Tests;

public sealed class NuGetConsumerIntegrationTests
{
    [Fact]
    public void PackedNuGet_ConsumerWithSourceGeneratedDto_Builds()
    {
        RunPackedConsumer(publishAot: false);
    }

    [Fact]
    public void PackedNuGet_SourceGeneratedDto_PublishesNativeAot()
    {
#if NET8_0
        RunPackedConsumer(publishAot: true);
#endif
    }

    private static void RunPackedConsumer(bool publishAot)
    {
        var repo = FindRepositoryRoot();
        var dotnet = Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "/usr/local/share/dotnet/dotnet";
        if (!File.Exists(dotnet)) dotnet = "dotnet";
        var root = Path.Combine(Path.GetTempPath(), "mieio-consumer-" + Guid.NewGuid().ToString("N"));
        var feed = Path.Combine(root, "feed");
        var consumer = Path.Combine(root, "consumer");
        Directory.CreateDirectory(feed);
        Directory.CreateDirectory(consumer);
        try
        {
            var packageVersion = "1.0.0-consumer-" + Guid.NewGuid().ToString("N");
            Run(dotnet, repo, "build src/Magicodes.IE.IO.SourceGenerator/Magicodes.IE.IO.SourceGenerator.csproj -c Debug -t:Rebuild -m:1 --no-restore");
            Run(dotnet, repo, $"pack src/Magicodes.IE.IO/Magicodes.IE.IO.csproj -c Debug --no-restore -m:1 -p:PackageVersion={packageVersion} -o \"{feed}\"");
            var package = Directory.GetFiles(feed, "Magicodes.IE.IO.*.nupkg").Single(p => !p.EndsWith(".symbols.nupkg", StringComparison.OrdinalIgnoreCase));
            File.WriteAllText(Path.Combine(consumer, "NuGet.config"), $"<configuration><packageSources><clear /><add key=\"local\" value=\"{feed}\" /><add key=\"nuget.org\" value=\"https://api.nuget.org/v3/index.json\" /></packageSources></configuration>");
            var publishProperties = publishAot ? "<PublishAot>true</PublishAot><PublishTrimmed>true</PublishTrimmed>" : string.Empty;
            File.WriteAllText(Path.Combine(consumer, "Consumer.csproj"), $"<Project Sdk=\"Microsoft.NET.Sdk\"><PropertyGroup><OutputType>Exe</OutputType><TargetFramework>net8.0</TargetFramework><Nullable>enable</Nullable>{publishProperties}</PropertyGroup><ItemGroup><PackageReference Include=\"Magicodes.IE.IO\" Version=\"{packageVersion}\" /></ItemGroup></Project>");
            File.WriteAllText(Path.Combine(consumer, "Program.cs"), "using System.IO; using System.Linq; using Magicodes.IE.IO; namespace Consumer; public sealed class MoneyConverter : CellConverter<decimal> { public override bool Read(string cell, out decimal value) => decimal.TryParse(cell, out value); } [XlsxExportable] public sealed class ConsumerRow { public string Name { get; set; } = \"\"; public decimal Amount { get; set; } } public static class Program { public static void Main() { var metadata = XlsxGeneratedTypeMetadataRegistry.TryGet<ConsumerRow>(); if (metadata is null) throw new System.Exception(\"metadata:null\"); using var ms = new MemoryStream(); Xlsx.Write(ms, new[] { new ConsumerRow { Name = \"ok\", Amount = 12.5m } }); ms.Position = 0; var options = new XlsxReadOptions<ConsumerRow>().WithConverter(new MoneyConverter()); var row = Xlsx.Read<ConsumerRow>(ms, options).Single(); if (row.Name != \"ok\" || row.Amount != 12.5m) throw new System.Exception($\"{row.Name}|{row.Amount}\"); } }");
            var rid = RuntimeInformation.RuntimeIdentifier;
            Run(dotnet, consumer, $"restore --no-cache --configfile NuGet.config{(publishAot ? $" -r {rid}" : string.Empty)} -v:minimal");
            if (!publishAot)
            {
                Run(dotnet, consumer, "build --no-restore -v:minimal");
                return;
            }

            Run(dotnet, consumer, $"publish --no-restore -c Release -r {rid} --self-contained true -p:PublishAot=true -p:PublishTrimmed=true -v:minimal");
            var publishRoot = Path.Combine(consumer, "bin", "Release", "net8.0", rid, "publish");
            var executable = Directory.GetFiles(publishRoot, "*", SearchOption.TopDirectoryOnly)
                .SingleOrDefault(path =>
                    string.Equals(Path.GetFileName(path), OperatingSystem.IsWindows() ? "Consumer.exe" : "Consumer", StringComparison.Ordinal));
            if (executable is null)
                throw new FileNotFoundException($"NativeAOT executable was not found under {publishRoot}. Files: {string.Join(", ", Directory.Exists(publishRoot) ? Directory.GetFiles(publishRoot, "*", SearchOption.AllDirectories) : Array.Empty<string>())}");
            Run(executable, consumer, string.Empty);
        }
        finally
        {
            try { Directory.Delete(root, recursive: true); } catch { }
        }
    }

    private static void Run(string fileName, string workingDirectory, string arguments)
    {
        using var process = Process.Start(new ProcessStartInfo(fileName, arguments)
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        }) ?? throw new InvalidOperationException("Could not start dotnet.");
        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();
        process.WaitForExit();
        Task.WaitAll(outputTask, errorTask);
        Assert.True(process.ExitCode == 0, $"{fileName} {arguments}\n{outputTask.Result}\n{errorTask.Result}");
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, "src", "Magicodes.IE.IO")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("Magicodes.IE repository root not found.");
    }
}
