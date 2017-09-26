using System;
using System.Collections.Generic;
using System.Text;

namespace TianCheng.Excel
{
    /// <summary>
    /// 实体对象属性与Excel中的列关系声明特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = true, Inherited = true)]
    public class ExcelColumnAttribute : System.Attribute
    {
        #region 特性属性
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
        #endregion 特性属性

        #region 构造方法
        /// <summary>
        /// 未设置列位置时，将使用列表末列的随机位置
        /// </summary>
        /// <param name="title">标题</param>
        public ExcelColumnAttribute(string title)
        {
            Title = title;
            Index = 0;
            ColName = String.Empty;
        }
        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="title">标题</param>
        /// <param name="index">列的序号</param>
        public ExcelColumnAttribute(string title, int index)
        {
            Title = title;
            Index = index;
            if (ExcelColumnIndexTran.Instance.IndexDict.ContainsKey(index))
            {
                ColName = ExcelColumnIndexTran.Instance.IndexDict[index];
            }
        }
        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="title">标题</param>
        /// <param name="colName">列的名称</param>
        public ExcelColumnAttribute(string title, string colName)
        {
            Title = title;
            ColName = colName.ToUpper();
            if (ExcelColumnIndexTran.Instance.ColumnDict.ContainsKey(colName))
            {
                Index = ExcelColumnIndexTran.Instance.ColumnDict[colName];
            }
        }
        #endregion 构造方法
    }
}
