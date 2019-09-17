namespace Magicodes.ExporterAndImporter.Core.Extension
{
    /// <summary>
    /// 扩展帮助类
    /// </summary>
    public static class Extension
    {
        /// <summary>
        /// string IsNullOrWhiteSpace扩展
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsNullOrWhiteSpace(this string str)
        {
            return string.IsNullOrWhiteSpace(str);
        }
    }
}