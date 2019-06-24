using System;
using System.Reflection;

namespace TianCheng.Excel
{
    /// <summary>
    /// 对象属性的操作
    /// </summary>
    internal class ObjectProperty
    {

        private static readonly string IntType = typeof(int).FullName;
        private static readonly string StringType = typeof(string).FullName;
        private static readonly string BoolType = typeof(bool).FullName;

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
                int.TryParse(Convert.ToString(val), out int iv);
                property.SetValue(instance, iv);
                return;
            }
            else if (property.PropertyType.FullName == BoolType)
            {
                bool.TryParse(Convert.ToString(val), out bool test);
                property.SetValue(instance, test);
                return;
            }
        }
    }
}
