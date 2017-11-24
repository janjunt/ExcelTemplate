using ExcelTemplate.Descriptor.Models;
using ExcelTemplate.Utility;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelTemplate.Utility.Extensions;
using ExcelTemplate.Utility.Models;

namespace ExcelTemplate.Descriptor
{
    /// <summary>
    /// 描述构造器接口
    /// </summary>
    public interface IDescriptorBuilder
    {
        /// <summary>
        /// 根据文件路径，构造参数描述列表
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>参数描述列表</returns>
        IList<ParameterDescriptor> Build(string filePath);
    }

    /// <summary>
    /// 描述构造器
    /// </summary>
    public class DefaultDescriptorBuilder : IDescriptorBuilder
    {
        #region 常量
        public const string ParameterPattern = @"\{\{([^\}]+)\}\}";
        #endregion

        #region 公开方法

        /// <summary>
        /// 根据文件路径，构造参数描述列表
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>参数描述列表</returns>
        public IList<ParameterDescriptor> Build(string filePath)
        {
            var result = new List<ParameterDescriptor>();
            var workbook = NpoiUtility.LoadWorkbook(filePath);
            for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
            {
                result.AddRange(BuildDescriptors(workbook.GetSheetAt(sheetIndex)));
            }

            return result;
        }
        #endregion

        #region 内部方法
        /// <summary>
        /// 根据sheet对象，构造参数描述列表
        /// </summary>
        /// <param name="sheet">sheet对象</param>
        /// <returns>参数描述列表</returns>
        private IList<ParameterDescriptor> BuildDescriptors(ISheet sheet)
        {
            var sheetIndex = sheet.GetSheetIndex();
            var result = BuildDescriptors(sheet.SheetName, sheetIndex);

            for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }

                for (var columnIndex = row.FirstCellNum; columnIndex <= row.LastCellNum; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    if (cell == null)
                    {
                        continue;
                    }

                    var cellValue = cell.ToString();
                    result.AddRange(BuildDescriptors(cellValue, sheetIndex, columnIndex, rowIndex));
                }
            }

            return result;
        }

        /// <summary>
        /// 根据原始字符串值，构造参数描述列表
        /// </summary>
        /// <param name="originalValue">原始字符串值</param>
        /// <param name="sheetIndex">所在Sheet</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="rowIndex">行索引</param>
        /// <returns>参数描述列表</returns>
        private IList<ParameterDescriptor> BuildDescriptors(
            string originalValue, 
            int sheetIndex,
            int? columnIndex = null, 
            int? rowIndex = null)
        {
            var result = new List<ParameterDescriptor>();

            var matches = Regex.Matches(originalValue, ParameterPattern);
            foreach (Match match in matches)
            {
                if (match.Success)
                {
                    var descriptor = new ParameterDescriptor(match.Groups[1].Value,originalValue, match.Groups[1].Index, sheetIndex);
                    if (columnIndex != null && rowIndex != null)
                    {
                        descriptor.Location = ParameterLocation.Cell;
                        descriptor.CellLocation = new CellLocation(columnIndex.Value, rowIndex.Value);
                    }
                    else
                    {
                        descriptor.Location = ParameterLocation.SheetName;
                    }

                    result.Add(descriptor);
                }
            }

            return result;
        }
        #endregion
    }
}
