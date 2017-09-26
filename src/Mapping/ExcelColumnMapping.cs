using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace TianCheng.Excel
{
    /// <summary>
    /// Excel 数据与对象的映射关系
    /// </summary>
    internal class ExcelColumnMapping
    {
        /// <summary>
        /// Excel中的标题名称
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 对象中的属性名称
        /// </summary>
        public string PropertyName { get; set; }
        /// <summary>
        /// 在Excel中在第几列  数字1开始
        /// </summary>
        public int Index { get; set; }
        /// <summary>
        /// Excel中的列名称： 例如：A 、B 
        /// </summary>
        public string ColName { get; set; }

        /// <summary>
        /// 属性信息
        /// </summary>
        internal PropertyInfo Property { get; set; }
    }
}
