using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    /// PDF native 库环境诊断和探测。
    /// wkhtmltopdf 的 native 库在各平台上都可能因缺少系统依赖而加载失败：
    /// - Linux: 缺少 fontconfig、libxrender、libjpeg、字体等（最常见）
    /// - macOS: 缺少 Homebrew Qt 框架
    /// - Windows: Haukcode.WkHtmlToPdfDotNet 自带，通常无需额外安装
    /// 本类不会主动加载 native 库，只做诊断并给出针对性的安装建议。
    /// </summary>
    internal static class PdfNativeLibraryBootstrapper
    {
        private static readonly Lazy<PdfEnvironmentInfo> CachedResult =
            new Lazy<PdfEnvironmentInfo>(CheckEnvironmentCore);

        /// <summary>
        /// 诊断当前环境，返回完整的 native 库状态。结果会被缓存，不会重复扫描文件系统。
        /// </summary>
        internal static PdfEnvironmentInfo CheckEnvironment() => CachedResult.Value;

        private static PdfEnvironmentInfo CheckEnvironmentCore()
        {
            var rid = RuntimeInformation.RuntimeIdentifier;
            var nativePath = FindNativeLibraryPath(rid);
            var loadable = TryProbeNativeLibrary(nativePath, out var error);

            return new PdfEnvironmentInfo
            {
                Platform = rid,
                NativeLibraryPath = nativePath,
                NativeLibraryFileExists = nativePath != null,
                NativeLibraryLoadable = loadable,
                IsAvailable = loadable,
                ErrorDetail = error,
                InstallSuggestion = GetInstallSuggestion(rid)
            };
        }

        /// <summary>
        /// 查找 native 库文件路径。
        /// 搜索 runtimes/{rid}/native/ 和应用根目录。
        /// </summary>
        private static string FindNativeLibraryPath(string rid)
        {
            var dirs = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "runtimes", rid, "native"),
                AppContext.BaseDirectory
            };

            var names = GetNativeLibraryNames(rid);
            foreach (var dir in dirs)
            {
                if (!Directory.Exists(dir)) continue;
                foreach (var name in names)
                {
                    var path = Path.Combine(dir, name);
                    if (File.Exists(path)) return path;
                }
            }
            return null;
        }

        private static string[] GetNativeLibraryNames(string rid)
        {
            if (rid.StartsWith("win", StringComparison.OrdinalIgnoreCase))
                return new[] { "wkhtmltox.dll" };
            if (rid.StartsWith("osx", StringComparison.OrdinalIgnoreCase))
                return new[] { "libwkhtmltox-next.0.dylib", "libwkhtmltox.dylib", "libwkhtmltox-next.dylib" };
            return new[] { "libwkhtmltox.so" };
        }

        /// <summary>
        /// 用 NativeLibrary.TryLoad 探测 native 库是否可加载。不抛异常。
        /// </summary>
        private static bool TryProbeNativeLibrary(string knownPath, out string error)
        {
#if NET6_0_OR_GREATER
            if (knownPath != null && NativeLibrary.TryLoad(knownPath, out _))
            {
                error = null;
                return true;
            }

            if (NativeLibrary.TryLoad("wkhtmltox", out _))
            {
                error = null;
                return true;
            }

            error = knownPath != null
                ? $"Found native library at '{knownPath}' but it failed to load. Missing system dependencies (Qt, fontconfig, etc.)?"
                : "No native library file found. wkhtmltopdf needs to be installed.";
            return false;
#else
            error = null;
            return true;
#endif
        }

        private static string GetInstallSuggestion(string rid)
        {
            if (rid.StartsWith("osx", StringComparison.OrdinalIgnoreCase))
                return "brew install wkhtmltopdf  # macOS";

            // Alpine Linux (musl libc) - Docker 最常用的基础镜像之一
            if (rid.Contains("musl", StringComparison.OrdinalIgnoreCase))
                return "apk add --no-cache wkhtmltopdf fontconfig ttf-freefont ttf-dejavu  # Alpine\n" +
                       "Note: Alpine uses musl libc. If using pre-built glibc binaries, also run: apk add --no-cache libc6-compat";

            if (rid.StartsWith("linux", StringComparison.OrdinalIgnoreCase))
                return "apt-get install -y wkhtmltopdf libxrender1 libfontconfig1 libjpeg62-turbo fontconfig xfonts-75dpi  # Debian/Ubuntu\n" +
                       "dnf install -y wkhtmltopdf fontconfig libXrender libjpeg-turbo xorg-x11-fonts-75dpi  # Fedora/RHEL 9+";

            if (rid.StartsWith("win", StringComparison.OrdinalIgnoreCase))
                return "wkhtmltopdf native library should be bundled with Haukcode.WkHtmlToPdfDotNet.\n" +
                       "If missing, download from https://wkhtmltopdf.org/downloads.html";

            return "Install wkhtmltopdf for your platform: https://wkhtmltopdf.org/downloads.html";
        }
    }

    public interface IPdfNativeLibraryService
    {
        PdfEnvironmentInfo CheckEnvironment();
    }

    public sealed class PdfNativeLibraryService : IPdfNativeLibraryService
    {
        public PdfEnvironmentInfo CheckEnvironment() => PdfNativeLibraryBootstrapper.CheckEnvironment();
    }
}
