using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace TianCheng.Excel
{
    /// <summary>
    /// Excel操作 http://blog.jobbole.com/92150/
    /// </summary>
    public class ExcelHelper
    {
        #region 数据导出
        /// <summary>
        /// 数据导出
        /// </summary>
        /// <returns></returns>
        static public string Export<T>(List<T> data, string excelFile)
        {
            //设置文件信息
            string dir = System.IO.Path.GetDirectoryName(excelFile);
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            FileInfo file = new FileInfo(excelFile);
            try
            {
                if (file.Exists)
                {
                    file.Delete();
                    file = new FileInfo(excelFile);
                }
            }
            catch (Exception)
            {
                throw new Exception("文件被占用，无法操作。");
            }

            //导出配置信息
            ExcelSheetMapping mapping = ExcelSheetMappingDict.Instance[typeof(T)];
            if (mapping == null)
            {
                throw new Exception("无法找到对象导出Excel的映射关系");
            }

            using (ExcelPackage package = new ExcelPackage(file))
            {
                //新增一个Sheet页
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(mapping.SheetName);

                int startRow = 1;
                //设置Sheet标题
                if (mapping.HasTitle)
                {
                    foreach (var map in mapping.ColumnMapping)
                    {
                        worksheet.Cells[startRow, map.Index].Value = map.Title;
                    }
                    startRow++;
                }
                //逐行设置数据
                foreach (T item in data)
                {
                    foreach (var map in mapping.ColumnMapping)
                    {

                        object val = map.Property.GetValue(item);
                        worksheet.Cells[startRow, map.Index].Value = val;
                    }
                    startRow++;
                }
                package.Save();
            }
            return file.FullName;

        }
        #endregion

        #region 数据导入
        /// <summary>
        /// 获取导入的对象信息
        /// </summary>
        /// <typeparam name="T">对象类型</typeparam>
        /// <param name="excelFile">Excel文件位置</param>
        /// <returns></returns>
        static public List<T> Import<T>(string excelFile) where T : new()
        {
            ExcelSheetMapping mapping = ExcelSheetMappingDict.Instance[typeof(T)];

            return Import<T>(excelFile, mapping);
        }


        /// <summary>
        /// 获取导入的对象信息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="excelFile"></param>
        /// <param name="mapping"></param>
        /// <returns></returns>
        static private List<T> Import<T>(string excelFile, ExcelSheetMapping mapping) where T : new()
        {
            FileInfo file = new FileInfo(excelFile);
            if (!file.Exists)
            {
                throw new Exception("无法找到要导入的Excel文件。");
            }
            if (mapping == null)
            {
                throw new Exception("对象与Excel的映射关系没有设置，请在对象上增加 ExcelSheetAttribute 特性");
            }

            try
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    StringBuilder sb = new StringBuilder();
                    //获取Sheet信息,如果按Sheet名称取不到，就取第一个Sheet页
                    ExcelWorksheet worksheet = null;
                    string sheetName = (mapping == null || String.IsNullOrEmpty(mapping.SheetName)) ? ExcelSheetAttribute.DefaultSheetName : mapping.SheetName;
                    if (package.Workbook == null || package.Workbook.Worksheets == null)
                    {
                        throw new Exception("无法读取Excel文件中的Sheet页信息");
                    }
                    worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                    {
                        worksheet = package.Workbook.Worksheets[1];
                    }

                    int rowCount = worksheet.Dimension.Rows;
                    int ColCount = worksheet.Dimension.Columns;
                    List<T> result = new List<T>();
                    TypeInfo type = typeof(T).GetTypeInfo();

                    //逐行获取数据
                    int startRow = mapping.HasTitle ? 2 : 1;
                    for (; startRow <= rowCount; startRow++)
                    {
                        T t = new T();
                        foreach (var map in mapping.ColumnMapping)
                        {
                            ObjectProperty.Set(t, map.Property, worksheet.Cells[startRow, map.Index].Value);
                        }
                        SetRowIndex(t, type, startRow);
                        result.Add(t);
                    }

                    return result;
                }
            }
            catch (Exception)
            {
                throw new Exception("Excel 文件格式不正确，或数据太大无法读取文件中的数据。建议数据不要超过5千行");
            }
        }

        static private void SetRowIndex(object instance, TypeInfo typeInfo, int rowIndex)
        {
            PropertyInfo pi = typeInfo.GetProperty("RowIndex");
            if (pi != null)
            {
                ObjectProperty.Set(instance, pi, rowIndex);
            }
        }
        #endregion 数据导入
    }
}
