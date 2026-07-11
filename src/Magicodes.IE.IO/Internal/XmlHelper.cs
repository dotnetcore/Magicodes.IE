
using System;

namespace Magicodes.IE.IO
{
    internal static class XmlHelper
    {
        internal static string EscapeXmlAttr(string s) => EscapeXmlText(s, isAttribute: true);

        internal static string EscapeXmlText(string s, bool isAttribute = false)
        {
            if (string.IsNullOrEmpty(s)) return s;

            int extra = 0;
            bool needsRewrite = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (char.IsHighSurrogate(c))
                {
                    if (i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                    {
                        i++;
                        continue;
                    }
                    needsRewrite = true;
                    continue;
                }
                if (char.IsLowSurrogate(c))
                {
                    needsRewrite = true;
                    continue;
                }
                extra += c switch
                {
                    '&' => 4,
                    '<' => 3,
                    '>' => 3,
                    '"' => isAttribute ? 5 : 0,
                    '\'' => isAttribute ? 5 : 0,
                    _ => 0,
                };

                if (IsIllegalXmlChar(c))
                {
                    needsRewrite = true;
                }
                else if (c is '&' or '<' or '>' or '"' or '\'')
                {
                    needsRewrite = true;
                }
            }

            if (!needsRewrite)
                return s;

            var chars = new char[s.Length + extra];
            int p = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (char.IsHighSurrogate(c) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                {
                    chars[p++] = c;
                    chars[p++] = s[++i];
                    continue;
                }
                if (char.IsSurrogate(c))
                {
                    chars[p++] = '\uFFFD';
                    continue;
                }
                switch (c)
                {
                    case '&':
                        chars[p++] = '&'; chars[p++] = 'a'; chars[p++] = 'm'; chars[p++] = 'p'; chars[p++] = ';';
                        break;
                    case '<':
                        chars[p++] = '&'; chars[p++] = 'l'; chars[p++] = 't'; chars[p++] = ';';
                        break;
                    case '>':
                        chars[p++] = '&'; chars[p++] = 'g'; chars[p++] = 't'; chars[p++] = ';';
                        break;
                    case '"':
                        if (isAttribute)
                        {
                            chars[p++] = '&'; chars[p++] = 'q'; chars[p++] = 'u'; chars[p++] = 'o'; chars[p++] = 't'; chars[p++] = ';';
                        }
                        else
                        {
                            chars[p++] = '"';
                        }
                        break;
                    case '\'':
                        if (isAttribute)
                        {
                            chars[p++] = '&'; chars[p++] = 'a'; chars[p++] = 'p'; chars[p++] = 'o'; chars[p++] = 's'; chars[p++] = ';';
                        }
                        else
                        {
                            chars[p++] = '\'';
                        }
                        break;
                    default:
                        chars[p++] = IsIllegalXmlChar(c) ? '\uFFFD' : c;
                        break;
                }
            }

            return new string(chars, 0, p);
        }

        private static bool IsIllegalXmlChar(char c)
        {
            return (c < 0x20 && c != '\t' && c != '\n' && c != '\r') || c is '\uFFFE' or '\uFFFF';
        }
    }
}
