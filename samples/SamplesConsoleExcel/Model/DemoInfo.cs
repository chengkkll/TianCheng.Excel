using System;
using System.Collections.Generic;
using System.Text;
using TianCheng.Excel;

namespace SamplesConsoleExcel.Model
{
    /// <summary>
    /// Excel中数据结构
    /// </summary>
    /// <remarks>ExcelSheet 特性标明是一个和Excel文件对应的对象类型</remarks>
    [ExcelSheet]
    public class DemoInfo
    {
        /// <summary>
        /// 企业名称
        /// </summary>
        [ExcelColumn("企业名称", "A")]
        public string Name { get; set; }

        /// <summary>
        /// 地址
        /// </summary>
        [ExcelColumn("地址", "B")]        
        public string Address { get; set; }

        /// <summary>
        /// 电话号码
        /// </summary>
        [ExcelColumn("电话号码", "C")]
        public string Telephone { get; set; }

    }
}
