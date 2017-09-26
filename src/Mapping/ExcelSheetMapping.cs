using System;
using System.Collections.Generic;
using System.Text;

namespace TianCheng.Excel
{
    /// <summary>
    /// Sheet 映射关系
    /// </summary>
    internal class ExcelSheetMapping
    {
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
        public string TypeName { get; set; }
        /// <summary>
        /// 对象类型全名
        /// </summary>
        public string TypeFullName { get; set; }

        private List<ExcelColumnMapping> _ColumnMapping = new List<ExcelColumnMapping>();
        /// <summary>
        /// 列的关系
        /// </summary>
        public List<ExcelColumnMapping> ColumnMapping
        {
            get { return _ColumnMapping; }
            set { _ColumnMapping = value; }
        }

    }
}
