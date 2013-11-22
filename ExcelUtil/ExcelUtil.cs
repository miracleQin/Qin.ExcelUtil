using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Qin.ExcelUtil
{
    public static class ExcelUtil
    {
        public static void ListToExcel<T>(this List<T> list, string filePath) where T : class
        {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("the filePath is empty or null.");
            if (list == null || list.Count == 0) throw new ArgumentException("The list is null or Count of list's items is 0");
            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
            Aspose.Cells.Worksheet sheet = wb.Worksheets[0];
            Aspose.Cells.Cells cells = sheet.Cells;

            //构建表名
            var modeDisplayNames = (DisplayNameAttribute[])typeof(T).GetCustomAttributes(typeof(DisplayNameAttribute), true);
            if (modeDisplayNames != null && modeDisplayNames.Length > 0)
                sheet.Name = modeDisplayNames[0].DisplayName;
            else
                sheet.Name = typeof(T).Name;

            //构建列名
            var keyValues = GetDisplayNames<T>();
            Dictionary<string, int> index_tmp = new Dictionary<string, int>();
            int index = 0;
            foreach (var key in keyValues.Keys)
            {
                cells[0, index].PutValue(key);
                index_tmp.Add(keyValues[key], index);
                index++;
            }
            //填充数据
            index = 1;
            foreach (var item in list)
            {
                var pis = item.GetType().GetProperties();
                foreach (var pi in pis)
                {
                    foreach (var key in index_tmp.Keys)
                    {
                        if (pi.Name == key)
                        {
                            cells[index, index_tmp[key]].PutValue(pi.GetValue(item, null));
                            break;
                        }
                    }
                }
                index++;
            }
            wb.Save(filePath);

        }

        public static List<T> ExcelToList<T>(string filePath) where T : class,new()
        {
            List<T> result = new List<T>();
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("The filePath can't be Null or Empty!");
            if (System.IO.File.Exists(filePath) == false) throw new ArgumentException("The File is not exists!");
            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(filePath);
            Aspose.Cells.Worksheet ws = wb.Worksheets[0];
            Aspose.Cells.Cells cells = ws.Cells;
            var keyValues = GetDisplayNames<T>();
            List<string> columnTiles = new List<string>();
            Dictionary<string, int> index_tmp = new Dictionary<string, int>();
            for (int i = 0; i <= cells.MaxColumn; i++)
            {
                foreach (var kv in keyValues)
                {
                    if (cells[0, i].StringValue == kv.Key || cells[0, i].StringValue == kv.Value)
                    {
                        if (columnTiles.Contains(kv.Value) == false)
                        {
                            columnTiles.Add(kv.Value);
                            index_tmp.Add(kv.Value, i);
                        }
                        break;
                    }
                }
            }

            for (int i = 1; i <= cells.MaxRow; i++)
            {
                T model = new T();
                PropertyInfo[] pis = model.GetType().GetProperties();
                bool isAdd = false;

                foreach (var pi in pis)
                {
                    foreach (var index in index_tmp)
                    {
                        if (pi.Name == index.Key)
                        {
                            var value = cells[i, index.Value].Value;
                            pi.SetValueInModel<T>(model, value);
                            isAdd = true;
                            break;
                        }
                    }
                }
                if (isAdd) result.Add(model);
            }


            return result;
        }

        private static Dictionary<string, string> GetDisplayNames<T>() where T : class
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            PropertyInfo[] pis = typeof(T).GetProperties().OrderBy(s => s.Name).ToArray();
            foreach (PropertyInfo pi in pis)
            {
                string name = "", value = "";
                value = pi.Name;
                var displayname = ClassExtensions<T>.GetDisplayNameAttribute(pi.Name);
                var display = ClassExtensions<T>.GetDisplayAttribute(pi.Name);
                if (displayname != null) name = displayname.DisplayName;
                else if (display != null) name = display.Name;
                else name = pi.Name;
                result.Add(name, value ?? name);
            }

            return result;
        }
        /// <summary>
        /// 设置反省对象属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="pi"></param>
        /// <param name="model"></param>
        /// <param name="value"></param>
        private static void SetValueInModel<T>(this PropertyInfo pi, T model, object value)
        {
            object end_value = null;
            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                var nullable = pi.PropertyType;
                if (nullable == typeof(int?))
                {
                    end_value = value.ToIntOrNull();
                }
                else if (nullable == typeof(short?))
                {
                    end_value = value.ToShortOrNull();
                }
                else if (nullable == typeof(long?))
                {
                    end_value = value.ToLongOrNull();
                }
                else if (nullable == typeof(double?))
                {
                    end_value = value.ToDoubleOrNull();
                }
                else if (nullable == typeof(float?))
                {
                    end_value = value.ToFloatOrNull();
                }
                else if (nullable == typeof(DateTime?))
                {
                    end_value = value.ToDateTimeOrNull();
                }
                else if (nullable == typeof(bool?))
                {
                    end_value = value.ToBoolOrNull();
                }
                else if (nullable == typeof(decimal?))
                {
                    end_value = value.ToDecimalOrNull();
                }
                else
                {
                    end_value = value == null ? null : value.ToString();
                }
            }
            else
            {
                switch (pi.PropertyType.Name)
                {
                    case "decimal":
                    case "Decimal":
                        end_value = value.ToDecimal();
                        break;
                    case "Boolean":
                    case "bool":
                        end_value = value.ToBool();
                        break;
                    case "Guid":
                        end_value = value.ToGuid();
                        break;
                    case "short":
                    case "int16":
                        end_value = value.ToShort();
                        break;
                    case "float":
                        end_value = value.ToFloat();
                        break;
                    case "DateTime":
                        end_value = value.ToDateTime();
                        break;

                    case "int32":
                    case "int":
                        end_value = value.ToInt();
                        break;

                    case "long":
                    case "int64":
                        end_value = value.ToLong();
                        break;

                    case "Double":
                    case "double":
                        end_value = value.ToDouble();
                        break;

                    default:
                        end_value = string.IsNullOrEmpty(value + "") ? null : value.ToString();
                        break;
                }
            }
            pi.SetValue(model, end_value, null);
        }

        #region 基础类型转换

        /* 可空类型，参数为空或者转换出错，返回值将是 null */
        #region 可空类型

        /// <summary>
        /// 将制定对象的值转换为 int? 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 null
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static int? ToIntOrNull(this object source)
        {
            int? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToInt32(source);
            }
            catch (Exception ex)
            {

            }
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 short? 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 null
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static short? ToShortOrNull(this object source)
        {
            short? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToInt16(source);
            }
            catch (Exception ex)
            {

            }
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 long? 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 null
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static long? ToLongOrNull(this object source)
        {
            long? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToInt64(source);
            }
            catch (Exception ex)
            {

            }
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 decimal? 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 null
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static decimal? ToDecimalOrNull(this object source)
        {
            decimal? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToDecimal(source);
            }
            catch (Exception ex)
            {

            }
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 double? 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 null
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static double? ToDoubleOrNull(this object source)
        {
            double? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToDouble(source);
            }
            catch (Exception ex)
            {

            }
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 float? 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 null
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static float? ToFloatOrNull(this object source)
        {
            float? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToSingle(source);
            }
            catch (Exception ex)
            {

            }
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 DateTime? 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 null
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static DateTime? ToDateTimeOrNull(this object source)
        {
            DateTime? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToDateTime(source);
            }
            catch (Exception ex)
            {

            }
            return result;
        }

        private static bool? ToBoolOrNull(this object source)
        {
            bool? result = null;
            try
            {
                if (source != null)
                    result = Convert.ToBoolean(source);
            }
            catch (Exception ex) { }
            return result;
        }

        private static Guid? ToGuidOrNull(this object source)
        {
            Guid? result = null;
            try
            {
                if (source != null)
                    result = new Guid(source.ToString());
            }
            catch (Exception ex) { }

            return result;
        }

        #endregion

        /*非可空类型 若值转换错误将用默认值代替。*/
        #region 非可空类型

        /// <summary>
        /// 将制定对象的值转换为 int 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 0
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static int ToInt(this object source)
        {
            int result = 0;
            if (source != null)
                int.TryParse(source.ToString(), out result);
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 short 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 0
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static short ToShort(this object source)
        {
            short result = 0;
            if (source != null)
                short.TryParse(source.ToString(), out result);
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 long 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 0
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static long ToLong(this object source)
        {
            long result = 0;
            if (source != null)
                long.TryParse(source.ToString(), out result);
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 decimal 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 0
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static decimal ToDecimal(this object source)
        {
            decimal result = 0;
            if (source != null)
                decimal.TryParse(source.ToString(), out result);
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 double 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 0
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static double ToDouble(this object source)
        {
            double result = 0;
            if (source != null)
                double.TryParse(source.ToString(), out result);
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 float 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 0
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static float ToFloat(this object source)
        {
            float result = 0;
            if (source != null)
                float.TryParse(source.ToString(), out result);
            return result;
        }
        /// <summary>
        /// 将制定对象的值转换为 DateTime 类型，如果传入值为空或者转换发生异常，
        /// 返回结果将是 new DateTimt()
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static DateTime ToDateTime(this object source)
        {
            DateTime result = DateTime.Now;
            if (source != null)
                DateTime.TryParse(source.ToString(), out result);
            return result;
        }
        /// <summary>
        /// 将类型转为 bool 
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        private static bool ToBool(this object source)
        {
            if (source != null && source.ToString().Equals("true", StringComparison.CurrentCultureIgnoreCase))
                return true;
            return false;
        }
        /// <summary>
        /// 将字符串转换成GUID，出错则为Guid.NewGuid()
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static Guid ToGuid(this object str)
        {
            Guid result = ToGuidOrNull(str).GetValueOrDefault(Guid.NewGuid());
            return result;
        }

        #endregion
        #endregion
    }

    #region 私有类
    /// <summary>
    /// 获取属性标注
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public static class ClassExtensions<T> where T : class
    {
        /**/
        /// <summary>
        /// 返回字段的Disaplay[name=""]等这些
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static DisplayAttribute GetDisplayAttribute(string propertyName)
        {
            return GetDisplayAttribute<DisplayAttribute>(propertyName, false);
        }

        /// <summary>
        /// 返回字段的DisaplayName等这些
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static DisplayNameAttribute GetDisplayNameAttribute(string propertyName)
        {
            return GetDisplayAttribute<DisplayNameAttribute>(propertyName, true);
        }
        /// <summary>
        /// 通用，返回类的字段的各种属性
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="propertyName"></param>
        /// <param name="attribute"></param>
        /// <returns></returns>
        public static TSource GetDisplayAttribute<TSource>(string propertyName, bool canNull) where TSource : class, new()
        {
            PropertyInfo propertyInfo = typeof(T).GetProperties().FirstOrDefault(s => s.Name == propertyName);
            if (propertyInfo == null)
                if (canNull) return null;
                else return new TSource();
            else
            {
                object[] attributes = propertyInfo.GetCustomAttributes(typeof(TSource), true);
                if (attributes.Length > 0)
                    return (TSource)attributes[0];
                else
                    if (canNull) return null;
                    else return new TSource();
            }
        }
    }

    #endregion
}
