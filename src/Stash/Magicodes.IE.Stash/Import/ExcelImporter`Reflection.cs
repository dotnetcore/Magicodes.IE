using CSScriptLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Magicodes.IE.Stash.Import
{
    public partial class ExcelImporter
    {
        /// <summary>
        /// 查询类型名字对应的ClrType
        /// </summary>
        /// <returns></returns>
        public static Func<string, Type> FindType = (typeName) =>
        {

            //先尝试直接获取类型
            var type = Type.GetType(typeName);

            // 这里要通过反映整个系统来获取对应的类型
            if (type == null)
            {
                var ass = AppDomain.CurrentDomain.GetAssemblies().ToList();

                List<Type> types = new();
                foreach (var item in ass)
                {
                    try
                    {
                        var ts = item.DefinedTypes.ToList();

                        foreach (var v in item.DefinedTypes)
                        {
                            try
                            {
                                types.Add(v.AsType());
                            }
                            catch (Exception ex)
                            {
                                //Console.WriteLine(ex);
                            }
                        }
                    }
                    catch (Exception aex)
                    {
                        //Console.WriteLine(aex);
                    }
                }

                type = types.Find(p => p.Name == typeName || p.FullName == typeName);
            }

            return type;
        };
    }
}
