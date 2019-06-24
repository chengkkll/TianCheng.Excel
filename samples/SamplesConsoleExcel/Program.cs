using SamplesConsoleExcel.Model;
using System;
using System.Collections.Generic;

namespace SamplesConsoleExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // 读取Excel文件中的内容到一个对象列表中
            List<DemoInfo> result = TianCheng.Excel.ExcelHelper.Import<DemoInfo>("../../../Files/Test.xlsx");

            Console.WriteLine($"共读取出{result.Count}条信息");
            foreach (var item in result)
                Console.WriteLine($"名称：{item.Name}\t地址：{item.Address}\t电话：{item.Telephone}");


            // 将对象导出到Excel文件中，如果文件存在，会覆盖。
            string file = "Export/export.xlsx";
            TianCheng.Excel.ExcelHelper.Export<DemoInfo>(result,file);
            Console.WriteLine($"文件已成功导出到:{file}");

            Console.Read();
        }
    }
}
