using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.StashTests.Extensions
{
    public static class Extensions
    {
        internal static JsonSerializerSettings JsonSerializerSettings => new()
        {
            //不用缩进美化代码
            Formatting = Formatting.None,
            //忽略循环依赖
            ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
            //忽略null值
            //NullValueHandling = NullValueHandling.Ignore,
            //不将首字母大写转小写
            ContractResolver = new Newtonsoft.Json.Serialization.DefaultContractResolver(),
            //日期格式化
            DateFormatString = "yyyy-MM-dd HH:mm:ss"
        };

        /// <summary>
        /// 将对象序列化为Json字符串
        /// </summary>
        /// <param name="obj">要序列化的对象</param>
        /// <param name="indented">指示是否格式化输出字符串</param>
        /// <param name="ignoreNullValue">是否忽略为null的属性(不序列化)</param>
        /// <returns></returns>
        public static string ToJsonString(this object obj, bool indented = false, bool ignoreNullValue = false)
        {
            if (obj != null)
            {
                var config = JsonSerializerSettings;
                if (indented)
                {
                    config.Formatting = Formatting.Indented;
                }
                if (ignoreNullValue)
                {
                    config.NullValueHandling = NullValueHandling.Ignore;
                }
                return JsonConvert.SerializeObject(obj, config);
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 将字符串反序列化为指定类型的实例
        /// </summary>
        /// <typeparam name="T">目标类型</typeparam>
        /// <param name="str">源字符串</param>
        /// <returns></returns>
        public static T ToJsonModel<T>(this string str) => JsonConvert.DeserializeObject<T>(str, JsonSerializerSettings);

    }
}
