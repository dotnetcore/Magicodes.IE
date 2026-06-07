using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
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
        /// 静态构造函数：首次访问类型时注册 DllImportResolver（不加载 native 库）。
        /// 确保 Haukcode 的 P/Invoke 能找到我们提供的 arm64 dylib。
        /// </summary>
        static PdfNativeLibraryBootstrapper()
        {
#if NET5_0_OR_GREATER
            RegisterDllImportResolver();
#endif
        }

        /// <summary>
        /// 诊断当前环境，返回完整的 native 库状态。结果会被缓存，不会重复扫描文件系统。
        /// </summary>
        internal static PdfEnvironmentInfo CheckEnvironment() => CachedResult.Value;

        private static PdfEnvironmentInfo CheckEnvironmentCore()
        {
            var rid = GetCurrentRuntimeIdentifier();
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
        /// 获取当前运行时标识符（跨框架兼容）。
        /// </summary>
        private static string GetCurrentRuntimeIdentifier()
        {
#if NET5_0_OR_GREATER
            return RuntimeInformation.RuntimeIdentifier;
#else
            var arch = RuntimeInformation.OSArchitecture.ToString().ToLowerInvariant();
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                return "win-" + arch;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                return "osx-" + arch;
            return "linux-" + arch;
#endif
        }

        /// <summary>
        /// 查找 native 库文件路径。
        /// 搜索 runtimes/{rid}/native/ 和应用根目录。
        /// 在 arm64 平台上如果找不到，会自动 fallback 到 x64 版本（通过 Rosetta 2 运行）。
        /// </summary>
        private static string FindNativeLibraryPath(string rid)
        {
            var baseDir = AppContext.BaseDirectory;
            var candidateRids = new List<string> { rid };

            // arm64 平台 fallback：wkhtmltopdf 等 native 库通常没有 arm64 版本，
            // x64 版本可通过 Rosetta 2 (macOS) 或兼容层 (Linux) 运行。
            if (rid.EndsWith("-arm64", StringComparison.OrdinalIgnoreCase))
            {
                var x64Rid = rid.Substring(0, rid.Length - "arm64".Length) + "x64";
                candidateRids.Add(x64Rid);
            }

            foreach (var candidateRid in candidateRids)
            {
                var dirs = new[]
                {
                    Path.Combine(baseDir, "runtimes", candidateRid, "native"),
                    baseDir
                };

                var names = GetNativeLibraryNames(candidateRid);
                foreach (var dir in dirs)
                {
                    if (!Directory.Exists(dir)) continue;
                    foreach (var name in names)
                    {
                        var path = Path.Combine(dir, name);
                        if (File.Exists(path)) return path;
                    }
                }
            }
            return null;
        }

        private static string[] GetNativeLibraryNames(string rid)
        {
            if (rid.StartsWith("win", StringComparison.OrdinalIgnoreCase))
                return new[] { "wkhtmltox.dll" };
            if (rid.StartsWith("osx", StringComparison.OrdinalIgnoreCase))
                return new[] { "libwkhtmltox.dylib" };
            return new[] { "libwkhtmltox.so" };
        }

        /// <summary>
        /// 用 NativeLibrary.TryLoad 探测 native 库是否可加载。不抛异常。
        /// 如果加载成功，注册 DllImportResolver 以便 Haukcode.WkHtmlToPdfDotNet
        /// 的 P/Invoke 能找到同一个已加载的库句柄（特别是 macOS arm64 场景，
        /// Haukcode 包不包含 osx-arm64 的 native 路径）。
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

#if NET5_0_OR_GREATER
        private static bool _resolverRegistered;

        /// <summary>
        /// 注册 DllImportResolver，将 "wkhtmltox" 延迟加载到 Haukcode 程序集。
        /// 不预加载库（避免 cocoa 插件在模块初始化时崩溃），
        /// 而是在 P/Invoke 首次调用时按需加载。
        /// </summary>
        private static void RegisterDllImportResolver()
        {
            if (_resolverRegistered) return;
            _resolverRegistered = true;

            try
            {
                NativeLibrary.SetDllImportResolver(
                    typeof(WkHtmlToPdfDotNet.BasicConverter).Assembly,
                    (libraryName, assembly, searchPath) =>
                    {
                        if (libraryName != "wkhtmltox")
                            return IntPtr.Zero;

                        // 延迟加载：首次 P/Invoke 调用时才加载 native 库
                        var rid = GetCurrentRuntimeIdentifier();
                        var path = FindNativeLibraryPath(rid);
                        if (path != null && NativeLibrary.TryLoad(path, out var handle))
                            return handle;
                        if (NativeLibrary.TryLoad("wkhtmltox", out handle))
                            return handle;
                        return IntPtr.Zero;
                    });
            }
            catch
            {
                // Already registered or not supported; ignore.
            }
        }
#endif

        private static string GetInstallSuggestion(string rid)
        {
            if (rid.StartsWith("osx", StringComparison.OrdinalIgnoreCase))
                return "The native library (libwkhtmltox.dylib) is bundled with Magicodes.IE.Pdf.\n" +
                       "No additional installation is required on macOS.";

            // Alpine Linux (musl libc) - Docker 最常用的基础镜像之一
            if (rid.IndexOf("musl", StringComparison.OrdinalIgnoreCase) >= 0)
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
