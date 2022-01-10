using System;
using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    public static class StringExtension
    {
        /// <summary>
        ///     判断指定的字符串是null、空还是空白字符
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsNullOrWhiteSpace(this string str)
        {
            return string.IsNullOrWhiteSpace(str);
        }

        private const char Base64Padding = '=';

        private static readonly HashSet<char> base64Table =
            new HashSet<char>{  'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O',
                'P','Q','R','S','T','U','V','W','X','Y','Z','a','b','c','d',
                'e','f','g','h','i','j','k','l','m','n','o','p','q','r','s',
                't','u','v','w','x','y','z','0','1','2','3','4','5','6','7',
                '8','9','+','/' };

        public static bool IsBase64String(this string value)
        {
#if NETSTANDARD2_1
            value = value.Replace("\r", string.Empty, StringComparison.Ordinal).Replace("\n", string.Empty, StringComparison.Ordinal);
#else
            value = value.Replace("\r", string.Empty).Replace("\n", string.Empty);
#endif

            if (value.Length == 0 || (value.Length % 4) != 0)
            {
                return false;
            }

            var lengthNoPadding = value.Length;
            value = value.TrimEnd(Base64Padding);
            var lengthPadding = value.Length;

            if ((lengthNoPadding - lengthPadding) > 2)
            {
                return false;
            }

            foreach (char c in value)
            {
                if (!base64Table.Contains(c))
                {
                    return false;
                }
            }

            return true;
        }

        public static bool IsBase64StringValid(this string value)
        {
            if (value == null)
            {
                return false;
            }

            return IsBase64String(value);
        }



    }
}
