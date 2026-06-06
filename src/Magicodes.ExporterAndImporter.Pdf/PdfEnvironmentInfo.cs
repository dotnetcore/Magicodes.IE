using System.Runtime.InteropServices;
using System.Text;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    /// PDF 导出环境诊断信息。
    /// 参考 SkiaSharp.SKImageInfo 和 PuppeteerSharp.BrowserFetcher 设计模式。
    /// 使用 <see cref="Check"/> 获取当前环境的完整诊断。
    /// </summary>
    /// <example>
    /// <code>
    /// var info = PdfEnvironmentInfo.Check();
    /// if (!info.IsAvailable)
    /// {
    ///     Console.WriteLine(info); // 输出：[FAIL] osx-arm64: wkhtmltopdf native library is NOT available
    ///                               //   Fix: brew install wkhtmltopdf
    /// }
    /// </code>
    /// </example>
    public sealed class PdfEnvironmentInfo
    {
        /// <summary>
        /// 当前平台标识（如 "osx-arm64", "linux-x64", "win-x64"）。
        /// </summary>
        public string Platform { get; init; }

        /// <summary>
        /// wkhtmltopdf native 库是否可用。
        /// 为 true 表示 PDF 导出功能可以正常工作。
        /// </summary>
        public bool IsAvailable { get; init; }

        /// <summary>
        /// 找到的 native 库文件路径。
        /// 可能来自 runtimes/ 目录或系统路径。
        /// </summary>
        public string NativeLibraryPath { get; init; }

        /// <summary>
        /// native 库文件是否存在（不论是否能加载）。
        /// 文件存在但加载失败通常意味着缺少系统依赖（Qt、fontconfig 等）。
        /// </summary>
        public bool NativeLibraryFileExists { get; init; }

        /// <summary>
        /// native 库是否能成功加载（通过 NativeLibrary.TryLoad 探测）。
        /// </summary>
        public bool NativeLibraryLoadable { get; init; }

        /// <summary>
        /// 加载失败时的详细错误信息。
        /// </summary>
        public string ErrorDetail { get; init; }

        /// <summary>
        /// 针对当前平台的安装建议。
        /// 包含具体的安装命令。
        /// </summary>
        public string InstallSuggestion { get; init; }

        /// <summary>
        /// 检查当前环境并返回诊断信息。不会抛出异常。
        /// </summary>
        /// <returns>包含完整诊断信息的 <see cref="PdfEnvironmentInfo"/> 实例。</returns>
        public static PdfEnvironmentInfo Check()
        {
            return PdfNativeLibraryBootstrapper.CheckEnvironment();
        }

        /// <summary>
        /// 返回格式化的诊断报告。
        /// </summary>
        public override string ToString()
        {
            if (IsAvailable)
                return $"[OK] {Platform}: native library loaded from {NativeLibraryPath}";

            var sb = new StringBuilder();
            sb.AppendLine($"[FAIL] {Platform}: wkhtmltopdf native library is NOT available");
            if (!string.IsNullOrEmpty(NativeLibraryPath))
                sb.AppendLine($"  Found: {NativeLibraryPath}");
            if (!string.IsNullOrEmpty(ErrorDetail))
                sb.AppendLine($"  Error: {ErrorDetail}");
            if (!string.IsNullOrEmpty(InstallSuggestion))
                sb.AppendLine($"  Fix:   {InstallSuggestion}");
            return sb.ToString();
        }
    }
}
