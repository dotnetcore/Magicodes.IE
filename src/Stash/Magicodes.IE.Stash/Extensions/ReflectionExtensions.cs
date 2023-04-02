using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.Stash.Extensions
{
    /// <summary>
    /// 反向扩展
    /// </summary>
    internal static class ReflectionExtensions
    {
        internal static bool CanBeSet(this MemberInfo member) => member is PropertyInfo property ? property.CanWrite : !((FieldInfo)member).IsInitOnly;

        internal static bool IsPublic(this PropertyInfo propertyInfo) => (propertyInfo.GetGetMethod() ?? propertyInfo.GetSetMethod()) != null;

    }
}
