namespace TianCheng.Excel
{
    /// <summary>
    /// 导入的实体对象接口定义
    /// </summary>
    public interface IImportExcelModel
    {
        /// <summary>
        /// 导入时的数据行号
        /// </summary>
        int RowIndex { get; set; }
    }
}
