using System;
using System.Collections.Generic;

namespace TianCheng.Excel
{
    /// <summary>
    /// Excel内Sheet的Mapping字典
    /// </summary>
    internal class ExcelSheetMappingDict
    {
        #region  获取一个对象实例
        /// <summary>
        /// 
        /// </summary>
        static public ExcelSheetMappingDict Instance { get; } = new ExcelSheetMappingDict();

        #endregion 获取一个对象实例

        #region 初始化对象
        private readonly Dictionary<string, ExcelSheetMapping> SheetMapping;
        /// <summary>
        /// 构造方法
        /// </summary>
        private ExcelSheetMappingDict()
        {
            //获取全部的对象与Excel的映射关系，并存入字典中。
            SheetMapping = new Dictionary<string, ExcelSheetMapping>();
            foreach (var mapping in GetMappingByAttribute.GetSheetMapping())
            {
                if (!SheetMapping.ContainsKey(mapping.TypeFullName))
                {
                    SheetMapping.Add(mapping.TypeFullName, mapping);
                }
            }
        }
        #endregion 初始化对象

        /// <summary>
        /// 根据对象类型获取与Excel的映射关系
        /// </summary>
        /// <param name="typeIndex"></param>
        /// <returns></returns>
        internal ExcelSheetMapping this[Type typeIndex]
        {
            get
            {
                string index = typeIndex.FullName;
                if (SheetMapping == null || !SheetMapping.ContainsKey(index))
                {
                    return null;
                }

                return SheetMapping[index];
            }
        }
    }
}
