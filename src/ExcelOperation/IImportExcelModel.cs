using System;
using System.Collections.Generic;
using System.Text;

namespace TianCheng.Excel
{
    public interface IImportExcelModel
    {
        /// <summary>
        /// 导入时的数据行号
        /// </summary>
        int RowIndex { get; set; }
    }
}
