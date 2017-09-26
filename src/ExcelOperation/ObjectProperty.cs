using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace TianCheng.Excel
{
    /// <summary>
    /// 对象属性的操作
    /// </summary>
    internal class ObjectProperty
    {

        private static string IntType = typeof(int).FullName;
        private static string StringType = typeof(String).FullName;
        private static string BoolType = typeof(bool).FullName;

        /// <summary>
        /// 设置对象属性
        /// </summary>
        /// <param name="instance"></param>
        /// <param name="property"></param>
        /// <param name="val"></param>
        static public void Set(object instance, PropertyInfo property, object val)
        {
            if (property.PropertyType.FullName == StringType)
            {
                property.SetValue(instance, Convert.ToString(val).Trim());
                return;
            }
            else if (property.PropertyType.FullName == IntType)
            {
                int iv = 0;
                int.TryParse(Convert.ToString(val), out iv);
                property.SetValue(instance, iv);
                return;
            }
            else if (property.PropertyType.FullName == BoolType)
            {
                bool test = false;
                bool.TryParse(Convert.ToString(val), out test);
                property.SetValue(instance, test);
                return;
            }
        }
    }
}
