using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.Stash.Import
{
    internal class Contract
    {
        /// <summary>
        /// 默认注入的命名空间
        /// <para>TODO: 这里要改造一下,允许用户来管理自己需要默认注入的命名空间,以减少模型定义时的工作</para>
        /// </summary>
        public static readonly List<string> DefaultNameSpaces = new()
        {
            "System",
            "System.Text",
            "System.Linq",
            "System.Collections.Generic"
        };

        //TODO:将模板表中的定界符号提到到这里,以便用户自定义
    }
}
