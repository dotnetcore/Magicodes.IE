// ======================================================================
// 
//           filename : Extension.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

#if NETSTANDARD2_1
using System.Reflection;
using System.Reflection.Emit;
#endif

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    /// <summary>
    ///     扩展类
    /// </summary>
    public static class Extension
    {
#if NETSTANDARD2_1
        /// <summary>
        /// 将集合转成DataTable
        /// </summary>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IEnumerable<T> source)
        {
            var table = Cache<T>.SchemeFactory();

            foreach (var item in source)
            {
                var row = table.NewRow();
                Cache<T>.Fill(row, item);
                table.Rows.Add(row);
            }

            return table;
        }

        private static class Cache<T>
        {
            // ReSharper disable StaticMemberInGenericType
            private static readonly PropertyInfo[] PropertyInfos;
            // ReSharper restore StaticMemberInGenericType

            // ReSharper disable StaticMemberInGenericType
            public static readonly Func<DataTable> SchemeFactory;
            // ReSharper restore StaticMemberInGenericType
            public static readonly Action<DataRow, T> Fill;

            static Cache()
            {
                PropertyInfos =
                    typeof(T).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                SchemeFactory = GenerateSchemeFactory();
                Fill = GenerateFill();
            }

            private static Func<DataTable> GenerateSchemeFactory()
            {
                var dynamicMethod =
                    new DynamicMethod($"__Extensions__SchemeFactory__Of__{typeof(T).Name}", typeof(DataTable),
                        Type.EmptyTypes, typeof(T), true);

                var generator = dynamicMethod.GetILGenerator();
                // ReSharper disable AssignNullToNotNullAttribute
                generator.Emit(OpCodes.Newobj, typeof(DataTable).GetConstructor(Type.EmptyTypes));
                // ReSharper restore AssignNullToNotNullAttribute
                generator.Emit(OpCodes.Dup);

                // ReSharper disable PossibleNullReferenceException
                generator.Emit(OpCodes.Callvirt, typeof(DataTable).GetProperty("Columns").GetMethod);
                // ReSharper restore PossibleNullReferenceException

                foreach (var propertyInfo in PropertyInfos)
                {
                    generator.Emit(OpCodes.Dup);
                    generator.Emit(OpCodes.Ldstr, propertyInfo.Name);
                    generator.Emit(OpCodes.Ldtoken, (propertyInfo.PropertyType.IsGenericType) && (propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>)) ? propertyInfo.PropertyType.GetGenericArguments()[0] : propertyInfo.PropertyType);
                    // ReSharper disable AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Call,
                        typeof(Type).GetMethod("GetTypeFromHandle", new[] { typeof(RuntimeTypeHandle) }));
                    // ReSharper restore AssignNullToNotNullAttribute
                    // ReSharper disable AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Callvirt,
                        typeof(DataColumnCollection).GetMethod("Add", new[] { typeof(string), typeof(Type) }));
                    // ReSharper restore AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Pop);
                }

                generator.Emit(OpCodes.Pop);
                generator.Emit(OpCodes.Ret);

                return (Func<DataTable>)dynamicMethod.CreateDelegate(typeof(Func<DataTable>));
            }

            private static Action<DataRow, T> GenerateFill()
            {
                var dynamicMethod =
                    new DynamicMethod($"__Extensions__Fill__Of__{typeof(T).Name}", typeof(void),
                        new[] { typeof(DataRow), typeof(T) }, typeof(T), true);

                var generator = dynamicMethod.GetILGenerator();
                for (var i = 0; i < PropertyInfos.Length; i++)
                {
                    generator.Emit(OpCodes.Ldarg_0);
                    switch (i)
                    {
                        case 0:
                            generator.Emit(OpCodes.Ldc_I4_0);
                            break;
                        case 1:
                            generator.Emit(OpCodes.Ldc_I4_1);
                            break;
                        case 2:
                            generator.Emit(OpCodes.Ldc_I4_2);
                            break;
                        case 3:
                            generator.Emit(OpCodes.Ldc_I4_3);
                            break;
                        case 4:
                            generator.Emit(OpCodes.Ldc_I4_4);
                            break;
                        case 5:
                            generator.Emit(OpCodes.Ldc_I4_5);
                            break;
                        case 6:
                            generator.Emit(OpCodes.Ldc_I4_6);
                            break;
                        case 7:
                            generator.Emit(OpCodes.Ldc_I4_7);
                            break;
                        case 8:
                            generator.Emit(OpCodes.Ldc_I4_8);
                            break;
                        default:
                            if (i <= 127)
                            {
                                generator.Emit(OpCodes.Ldc_I4_S, (byte)i);
                            }
                            else
                            {
                                generator.Emit(OpCodes.Ldc_I4, i);
                            }
                            break;
                    }
                    generator.Emit(OpCodes.Ldarg_1);
                    generator.Emit(OpCodes.Callvirt, PropertyInfos[i].GetGetMethod(true));
                    if (PropertyInfos[i].PropertyType.IsValueType)
                    {
                        generator.Emit(OpCodes.Box, PropertyInfos[i].PropertyType);
                    }
                    // ReSharper disable AssignNullToNotNullAttribute
                    generator.Emit(OpCodes.Callvirt, typeof(DataRow).GetMethod("set_Item", new[] { typeof(int), typeof(object) }));
                    // ReSharper restore AssignNullToNotNullAttribute
                }
                generator.Emit(OpCodes.Ret);
                return (Action<DataRow, T>)dynamicMethod.CreateDelegate(typeof(Action<DataRow, T>));
            }

        }
#else

        /// <summary>
        ///     将集合转成DataTable
        /// </summary>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this ICollection<T> source)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p =>
                new DataColumn(p.Name,
                    (p.PropertyType.IsGenericType) && (p.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                        ? p.PropertyType.GetGenericArguments()[0]
                        : p.PropertyType)).ToArray());

            for (var i = 0; i < source.Count; i++)
            {
                var tempList = new ArrayList();
                foreach (var obj in props.Select(pi => pi.GetValue(source.ElementAt(i), null))) tempList.Add(obj);
                var array = tempList.ToArray();
                dt.LoadDataRow(array, true);
            }

            return dt;
        }

#endif
        /// <summary>
        ///     将DataTable转List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static IList<T> ToList<T>(this DataTable dt) where T : class
        {
            IList<T> list = new List<T>();
            foreach (DataRow dr in dt.Rows)
            {
                T t = Activator.CreateInstance<T>();
                var props = typeof(T).GetProperties();
                foreach (var pro in props)
                {
                    var tempName = pro.Name;
                    if (!dt.Columns.Contains(tempName)) continue;
                    if (!pro.CanWrite) continue;
                    var value = dr[tempName];
                    if (value != DBNull.Value)
                        pro.SetValue(t, value, null);
                }

                list.Add(t);
            }

            return list;
        }


        /// <summary>
        /// 将Bytes导出为Excel文件
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="fileName">文件路径</param>
        /// <returns></returns>
        public static ExportFileInfo ToExcelExportFileInfo(this byte[] bytes, string fileName)
        {
            fileName.CheckExcelFileName();
            File.WriteAllBytes(fileName, bytes);

            var file = new ExportFileInfo(fileName,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            return file;
        }

        /// <summary>
        /// 将Bytes导出为Csv文件
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="fileName">文件路径</param>
        /// <returns></returns>
        public static ExportFileInfo ToCsvExportFileInfo(this byte[] bytes, string fileName)
        {
            fileName.CheckCsvFileName();
            File.WriteAllBytes(fileName, bytes);

            var file = new ExportFileInfo(fileName,
                "text/csv");

            return file;
        }

        /// <summary>
        /// 检查文件名
        /// </summary>
        /// <param name="fileName"></param>
        public static void CheckExcelFileName(this string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException("文件名必须填写!", nameof(fileName));
            if (!Path.GetExtension(fileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("仅支持导出“.xlsx”，即不支持Excel97-2003!", nameof(fileName));
            }
        }

        /// <summary>
        /// 检查文件名
        /// </summary>
        /// <param name="fileName"></param>
        public static void CheckCsvFileName(this string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("文件名必须填写!", nameof(fileName));
            if (!Path.GetExtension(fileName).Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("仅支持导出“.csv”!", nameof(fileName));
            }
        }
        /// <summary>
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static Bitmap GetBitmapByUrl(string url)
        {
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            var wc = new System.Net.WebClient();
            return new Bitmap(wc.OpenRead(url));
        }

        /// <summary>
        /// 保存图片
        /// </summary>
        /// <param name="image"></param>
        /// <param name="path">path</param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static string Save(this Image image, string path, ImageFormat format)
        {
            using (var img = image)
            {
                img.Save(path, format);
            }
            return path;
        }
        /// <summary>
        ///     图片转base64
        /// </summary>
        /// <param name="image"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static string ToBase64String(this Image image, ImageFormat format)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, format);
                var arr = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(arr, 0, (int)ms.Length);
                ms.Close();
                return Convert.ToBase64String(arr);
            }
        }

        /// <summary>
        ///     base64转Bitmap
        /// </summary>
        /// <param name="base64String"></param>
        /// <returns></returns>
        public static Bitmap Base64StringToBitmap(this string base64String)
        {
            var bitmapData = Convert.FromBase64String(s: FixBase64ForImage(Image: base64String));
            var streamBitmap = new MemoryStream(buffer: bitmapData);
            var bitmap = new Bitmap(original: (Bitmap)Image.FromStream(stream: streamBitmap));
            return bitmap;
        }

        private static string FixBase64ForImage(string Image)
        {
            var sbText = new System.Text.StringBuilder(Image, Image.Length);
            sbText.Replace("\r\n", string.Empty);
            sbText.Replace(" ", string.Empty);
            return sbText.ToString();
        }

        /// <summary>
        ///     获取集合连续数据中最大的
        /// </summary>
        /// <param name="numList"></param>
        /// <returns></returns>
        public static int GetLargestContinuous(this List<int> numList)
        {
            for (int i = 0; i < numList.Count;)
            {
                if (numList.Count > i + 1 && numList[i] - numList[i + 1] == 1)
                {
                    //忽略
                }

                return numList[i];
            }

            return 0;
        }
    }
}