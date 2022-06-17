using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EasyOffice.Models.Word
{
    /// <summary>
    /// 规范模板数据类型，针对自定义模板文档，又有占位符字段又有数据列表数据时，对数据列表基类进行
    /// </summary>
    public abstract class BaseTableEle
    {
        public Dictionary<string, string> GetKeyValuePairs()
        {
            Dictionary<string, string> replacements = new Dictionary<string, string>();
            Type type = this.GetType();
            //列表对象只允许使用简单数据类型
            //PropertyInfo[] props = type.GetProperties().Where(x => x.PropertyType == typeof(string) || x.PropertyType == typeof(int) || x.PropertyType == typeof(double) || x.PropertyType == typeof(long) || x.PropertyType == typeof(decimal) || x.PropertyType.)
            //    .ToArray();
            PropertyInfo[] props = type.GetProperties().Where(x => x.PropertyType == typeof(string) || x.PropertyType.IsValueType).ToArray();

            foreach (PropertyInfo prop in props)
            {
                if (!prop.CanRead || !prop.CanWrite)
                    continue;

                var replacement = prop.GetValue(this)?.ToString();

                var placeholder = prop.IsDefined(typeof(PlaceholderAttribute))
                                ? prop.GetCustomAttribute<PlaceholderAttribute>().Placeholder.ToString()
                                : prop.Name;

                replacements.Add(placeholder, replacement);
            }

            return replacements;
        }
    }
}