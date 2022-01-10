// ======================================================================
// 
//           filename : HtmlExporter.cs
//           description :
// 
//           created by 雪雁 at  2019-09-26 14:59
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using RazorEngine.Compilation;
using RazorEngine.Compilation.ReferenceResolver;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Magicodes.ExporterAndImporter.Html
{
    /// <summary>
    /// 
    /// </summary>
    internal class ExternalAssemblyReferenceResolver : IReferenceResolver
    {
        private string[] _assembliesToLoad;
        public ExternalAssemblyReferenceResolver(params string[] assembliesToLoad)
        {
            _assembliesToLoad = assembliesToLoad;
        }

        public IEnumerable<CompilerReference> GetReferences(TypeContext context, IEnumerable<CompilerReference> includeAssemblies = null)
        {
            var results = new UseCurrentAssembliesReferenceResolver()
           .GetReferences(context, includeAssemblies)
           .ToList();
            //添加DataTable的引用
            results.Add(CompilerReference.From(typeof(DataTable).Assembly));
            if (_assembliesToLoad != null)
            {
                foreach (var item in _assembliesToLoad)
                {
                    results.Add(CompilerReference.From(item));
                }
            }
            return results;
        }
    }
}