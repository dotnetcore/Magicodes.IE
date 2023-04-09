using CSScriptLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Magicodes.IE.Stash.Import
{
    public partial class ExcelImporter
    {
        private Func<string, string> NormalizeNameSpace = (p) =>
        {
            string ret = p.Trim();

            //TODO: 如果用户将using的大小写搞错了,要不要自动帮他修复.这里先不管了
            if (p.StartsWith("using "))
            {
                ret = p;
            }
            else
            {
                ret = $"using {p}";
            }

            // 追加语句后的分号
            if (!ret.EndsWith(";"))
            {
                ret = $"{ret};";
            }
            return ret;
        };

        /// <summary>
        /// 合并命名空间
        /// </summary>
        /// <param name="nameSpaces"></param>
        /// <returns></returns>
        private List<string> MergeNameSpaces(List<string> nameSpaces)
        {
            var retList = nameSpaces
                .Select(p => string.IsNullOrWhiteSpace(p) ? null : NormalizeNameSpace(p))
                .ToList();

            //注入默认命名空间
            retList.AddRange(Contract.DefaultNameSpaces
                .Select(p => string.IsNullOrWhiteSpace(p) ? null : NormalizeNameSpace(p)));

            retList = retList.OrderBy(p => p).Distinct().Where(p => p != null).ToList();

            return retList!;
        }

        private (string FullCode, IFunc func) BuildCode(string code, List<string> nameSpaces)
        {
            var sb = new StringBuilder();

            MergeNameSpaces(nameSpaces)
                .ForEach(p => sb.AppendLine(p));//添加命令空间

            sb.AppendLine($"object Calc(dynamic p)");
            sb.AppendLine("{");
            sb.AppendLine(code);
            sb.AppendLine(";");
            sb.AppendLine("}");

            var _code = sb.ToString();

            Console.WriteLine(_code);

            var calc = CSScript.Evaluator.LoadMethod<IFunc>(_code);

            return (_code, calc);
        }

        /// <summary>
        /// 编译变量
        /// <para>编译定义中的变量,计算变量值,</para>
        /// </summary>
        private void BuidVariables()
        {
            // 检查变量名是否重复
            var _repeatedNames = MapDefinition.Variables.GroupBy(p => p.Name)
                .Where(p => p.Count() > 1)
                .Select(p => p.Key).ToList();

            if (_repeatedNames.Any())
            {
                throw new Exception($"变量名重复:{string.Join(",", _repeatedNames)}");
            }

            MapDefinition.Variables.ForEach(p =>
            {
                var (FullCode, func) = BuildCode(p.Code, MapDefinition.Namespaces);
                p.FullCode = FullCode;
                p.Func = func;
            });
        }

        /// <summary>
        /// 编译转换器
        /// </summary>
        private void BuildConverter()
        {
            MapDefinition.Maps.ForEach(map => map.Pipes.ForEach(pipe =>
            {
                var (FullCode, func) = BuildCode(pipe.Code, MapDefinition.Namespaces);
                pipe.FullCode = FullCode;
                pipe.Func = func;
            }));
        }

        /// <summary>
        /// 编译
        /// <para>编译转换器和变量代码,并且会计算变量的值,请勿将编译后的变量结果持久化</para>
        /// </summary>

        //TODO: 应该将编译步骤拆分开,做一个分析功能,用来检查模板中的语法是否正确.
        public void Build()
        {
            BuidVariables();
            BuildConverter();
        }
    }
}
