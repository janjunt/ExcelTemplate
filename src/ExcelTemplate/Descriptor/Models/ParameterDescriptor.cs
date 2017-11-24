using ExcelTemplate.Utility.Models;

namespace ExcelTemplate.Descriptor.Models
{
    /// <summary>
    /// 参数描述
    /// </summary>
    public class ParameterDescriptor
    {
        #region 属性
        /// <summary>
        /// 描述值
        /// </summary>
        public string Value { get; set; }
        /// <summary>
        /// 原始值
        /// </summary>
        public string OriginalValue { get; set; }
        /// <summary>
        /// 参数起始位置索引
        /// </summary>
        public int StartIndex { get; set; }
        /// <summary>
        /// 参数所在位置类型
        /// </summary>
        public ParameterLocation Location { get; set; }
        /// <summary>
        /// 所在Sheet
        /// </summary>
        public int SheetIndex { get; set; }
        /// <summary>
        /// 所在单元格
        /// </summary>
        public CellLocation CellLocation { get; set; }
        #endregion

        #region 构造函数
        /// <summary>
        /// 有参构造函数
        /// </summary>
        /// <param name="value">描述值</param>
        /// <param name="originalValue">原始值</param>
        /// <param name="startIndex">参数起始位置索引</param>
        /// <param name="sheetIndex">所在Sheet</param>
        public ParameterDescriptor(
            string value, 
            string originalValue, 
            int startIndex,
            int sheetIndex)
        {
            Value = value;
            OriginalValue = originalValue;
            StartIndex = startIndex;
            SheetIndex = sheetIndex;
        }
        #endregion
    }
}
