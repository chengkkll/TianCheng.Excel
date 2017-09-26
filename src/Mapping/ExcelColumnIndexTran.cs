using System;
using System.Collections.Generic;
using System.Text;

namespace TianCheng.Excel
{
    /// <summary>
    /// Excel列名与序号的转换
    /// </summary>
    internal class ExcelColumnIndexTran
    {
        #region  获取一个对象实例
        static private ExcelColumnIndexTran _Instance = new ExcelColumnIndexTran();

        static public ExcelColumnIndexTran Instance
        {
            get
            {
                return _Instance;
            }
        }
        #endregion 获取一个对象实例
        /// <summary>
        /// 按序号查询的字典
        /// </summary>
        public Dictionary<int, string> IndexDict { get; private set; }
        /// <summary>
        /// 按列查询的字典
        /// </summary>
        public Dictionary<string, int> ColumnDict { get; private set; }

        /// <summary>
        /// 构造方法
        /// </summary>
        private ExcelColumnIndexTran()
        {
            IndexDict = new Dictionary<int, string>();
            ColumnDict = new Dictionary<string, int>();
            //A-Z
            for (int i = 1; i <= 26; i++)
            {
                IndexDict.Add(i, Convert.ToChar(64 + i).ToString());
                ColumnDict.Add(Convert.ToChar(64 + i).ToString(), i);
            }
            //AA - AZ
            for (int i = 1; i <= 26; i++)
            {
                IndexDict.Add(26 + i, "A" + Convert.ToChar(64 + i).ToString());
                ColumnDict.Add("A" + Convert.ToChar(64 + i).ToString(), 26 + i);
            }
        }
    }
}
