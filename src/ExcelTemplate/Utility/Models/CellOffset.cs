namespace ExcelTemplate.Utility.Models
{
    /// <summary>
    /// 单元格偏移
    /// </summary>
    public class CellOffset
    {
        #region 属性
        /// <summary>
        /// 横向偏移量
        /// </summary>
        public int OffsetX { get; set; }
        /// <summary>
        /// 纵向偏移量
        /// </summary>
        public int OffsetY { get; set; }
        #endregion

        #region 构造函数

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="offsetX">横向偏移量</param>
        /// <param name="offsetY">纵向偏移量</param>
        public CellOffset(int offsetX = 0, int offsetY = 0)
        {
            OffsetX = offsetX;
            OffsetY = offsetY;
        }
        #endregion

        #region 运算符重载
        /// <summary>
        /// 重载+运算符
        /// </summary>
        /// <param name="offset1">单元格偏移1</param>
        /// <param name="offset2">单元格偏移2</param>
        /// <returns></returns>
        public static CellOffset operator +(CellOffset offset1, CellOffset offset2)
        {
            return new CellOffset(offset1.OffsetX + offset2.OffsetX, offset1.OffsetY + offset2.OffsetY);
        }
        #endregion
    }

    /// <summary>
    /// 单元格偏移扩展方法
    /// </summary>
    public static class CellOffsetExtensions
    {
        /// <summary>
        /// 复制单元格偏移
        /// </summary>
        /// <param name="source">源单元格偏移</param>
        /// <returns>复制的单元格偏移</returns>
        public static CellOffset Copy(this CellOffset source)
        {
            return new CellOffset(source.OffsetX, source.OffsetY);
        }
    }
}
