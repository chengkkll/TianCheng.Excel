# TianCheng.Excel

目标架构为：.NET Standard 2.0

封装了EPPlus，实现可以按对象读写Excel。

## 快速使用

### 定义对象与Excel的关系

`ExcelSheet`特性用于类上，设置对象与`Excel`的`Sheet`页的信息。`ExcelSheet`有两个属性，`sheetName`表示`sheet`页的名称，`sheetName`默认为`Sheet1`；`hasTitle`表示当前`sheete`页是否有标题信息。

`ExcelColumn`特性用于属性上，设置每一列的信息。`ExcelColumn`可以设置3个属性，`Title`表示列显示的标题信息；`Index`表示列所在的序号（数字）；`ColName`表示列所在的列号（字母，如A、B、C...AA、AB）

使用示例

  ```csharp
    /// <summary>
    /// Excel中数据结构
    /// </summary>
    /// <remarks>ExcelSheet 特性标明是一个和Excel文件对应的对象类型</remarks>
    [ExcelSheet]
    public class DemoInfo
    {
        /// <summary>
        /// 名称
        /// </summary>
        [ExcelColumn("名称", "A")]
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
  ```

### 读取Excel文件中的数据

ExcelHelper.Import可将指定文件中的数据转成对象列表。

示例

  ```csharp
      // 读取Excel文件中的内容到一个对象列表中
      List<DemoInfo> result = TianCheng.Excel.ExcelHelper.Import<DemoInfo>("Files/Test.xlsx");
  ```

### 写入Excel文件

ExcelHelper.Export 可以将数据列表写入指定的文件中

示例

  ```csharp
      // 读取Excel文件中的内容到一个对象列表中
      TianCheng.Excel.ExcelHelper.Export<DemoInfo>(result,"Export/export.xlsx");
  ```

### 完整调用例子

  ```csharp
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
  ```
