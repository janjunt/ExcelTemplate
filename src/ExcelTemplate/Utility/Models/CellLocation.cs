using ExcelTemplate.Descriptor.Models;

namespace ExcelTemplate.Utility.Models
{
    /// <summary>
    /// 单元格位置
    /// </summary>
    public class CellLocation
    {
        #region 属性

        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }
        #endregion

        #region 构造函数

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <param name="rowIndex">行索引</param>
        public CellLocation(int columnIndex, int rowIndex)
        {
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
        }
        #endregion

        #region 运算符重载
        /// <summary>
        /// 重载+运算符
        /// </summary>
        /// <param name="location">单元格位置</param>
        /// <param name="offset">单元格偏移</param>
        /// <returns>运算后的单元格位置</returns>
        public static CellLocation operator +(CellLocation location, CellOffset offset)
        {
            return new CellLocation(location.ColumnIndex + offset.OffsetX, location.RowIndex + offset.OffsetY);
        }
        #endregion
    }

    /// <summary>
    /// 单元格位置扩展方法
    /// </summary>
    public static class CellLocationExtensions
    {
        /// <summary>
        /// 复制单元格位置
        /// </summary>
        /// <param name="source">源单元格位置</param>
        /// <returns>复制的单元格位置</returns>
        public static CellLocation Copy(this CellLocation source)
        {
            return new CellLocation(source.ColumnIndex, source.RowIndex);
        }
    }
}
