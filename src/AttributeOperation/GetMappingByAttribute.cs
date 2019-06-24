using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace TianCheng.Excel
{
    /// <summary>
    /// Excel与对象映射关系处理
    /// </summary>
    internal class GetMappingByAttribute
    {
        /// <summary>
        /// 获取所有可用的Excel与对象的映射关系
        /// </summary>
        /// <returns></returns>
        static internal List<ExcelSheetMapping> GetSheetMapping()
        {
            List<Assembly> assemblyList = TianCheng.Model.AssemblyHelper.GetAssemblyList();
            List<ExcelSheetMapping> mappingList = new List<ExcelSheetMapping>();
            foreach (Assembly assembly in assemblyList)
            {
                mappingList.AddRange(GetSheetMapping(assembly));
            }
            return mappingList;
        }

        /// <summary>
        /// 在程序集对象的映射关系
        /// </summary>
        /// <param name="assembly"></param>
        /// <returns></returns>
        static private List<ExcelSheetMapping> GetSheetMapping(Assembly assembly)
        {
            List<ExcelSheetMapping> mappingList = new List<ExcelSheetMapping>();

            foreach (var type in assembly.GetTypes())
            {
                TypeInfo ti = type.GetTypeInfo();
                ExcelSheetAttribute attribute = ti.GetCustomAttribute<ExcelSheetAttribute>(false);    //false 不获取基类中的特性
                if (attribute == null)
                {
                    continue;   //如果类中不包含Excel导出导出特性跳过。
                }

                ExcelSheetMapping sheet = new ExcelSheetMapping
                {
                    TypeName = type.Name,
                    TypeFullName = type.FullName,
                    SheetName = attribute.SheetName,
                    HasTitle = attribute.HasTitle,
                    //根据特性设置每一个属性值的情况
                    ColumnMapping = GetColumnMapping(ti)
                };
                if (sheet.ColumnMapping.Count <= 0)
                {
                    continue;
                }

                mappingList.Add(sheet);
            }

            return mappingList;
        }


        /// <summary>
        /// 获取对象属性的映射关系
        /// </summary>
        /// <param name="typeInfo"></param>
        /// <returns></returns>
        static private List<ExcelColumnMapping> GetColumnMapping(TypeInfo typeInfo)
        {
            List<ExcelColumnMapping> mapping = new List<ExcelColumnMapping>();
            //循环设置每一个拥有导入导出特性的属性信息
            foreach (PropertyInfo prop in typeInfo.GetProperties())
            {
                try
                {
                    ExcelColumnAttribute attribute = prop.GetCustomAttribute<ExcelColumnAttribute>(true);
                    if (attribute == null)
                    {
                        continue;
                    }
                    ExcelColumnMapping column = new ExcelColumnMapping
                    {
                        ColName = attribute.ColName,
                        Index = attribute.Index,
                        Title = attribute.Title,
                        PropertyName = prop.Name,
                        Property = prop
                    };
                    mapping.Add(column);
                }
                catch (Exception)
                {
                    throw;
                }
            }

            //如果未设置Excel对应的列，设置为末列
            var emptyList = mapping.Where(e => e.Index == 0).ToList();
            if (emptyList.Count > 0)
            {
                int max = mapping.Max(e => e.Index);
                foreach (var item in emptyList)
                {
                    item.Index = ++max;
                    if (ExcelColumnIndexTran.Instance.IndexDict.ContainsKey(item.Index))
                    {
                        item.ColName = ExcelColumnIndexTran.Instance.IndexDict[item.Index];
                    }
                }
            }

            return mapping;
        }
    }
}
