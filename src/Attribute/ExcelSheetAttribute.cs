using System;
using System.Collections.Generic;
using System.Text;

namespace TianCheng.Excel
{
    /// <summary>
    /// 实体对象属性与Excel中的列关系声明特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
    public class ExcelSheetAttribute : System.Attribute
    {
        #region 特性属性
        /// <summary>
        /// Sheet名称
        /// </summary>
        public string SheetName { get; set; }
        /// <summary>
        /// 导入数据中是否包含标题
        /// </summary>
        public bool HasTitle { get; set; }

        /// <summary>
        /// 对象类型名称
        /// </summary>
        internal string TypeName { get; set; }
        /// <summary>
        /// 对象类型全名
        /// </summary>
        internal string TypeFullName { get; set; }
        #endregion 特性属性

        #region 构造方法

        /// <summary>
        /// 默认的Sheet名称
        /// </summary>
        private const string DefaultSheetName = "Sheet1";

        /// <summary>
        /// 构造方法
        /// </summary>
        public ExcelSheetAttribute()
        {
            SheetName = DefaultSheetName;
            HasTitle = true;
        }

        /// <summary>
        /// 构造方法
        /// </summary>_
        /// <param name="title">标题</param>
        /// <param name="index">列的序号</param>
        public ExcelSheetAttribute(string sheetName, bool hasTitle = true)
        {
            if (String.IsNullOrEmpty(sheetName))
            {
                sheetName = DefaultSheetName;
            }
            SheetName = sheetName;
            HasTitle = hasTitle;
        }
        #endregion 构造方法
    }
}
