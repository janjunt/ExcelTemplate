using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using ExcelTemplate.Utility.Models;

namespace ExcelTemplate.Utility.Extensions
{
    public static class SheetExtensions
    {
        public static int GetSheetIndex(this ISheet sheet)
        {
            if (sheet == null || sheet.Workbook == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            return sheet.Workbook.GetSheetIndex(sheet);
        }

        public static void SetSheetName(this ISheet sheet, string sheetName)
        {
            if (sheet == null || sheet.Workbook == null)
            {
                return;
            }

            sheet.Workbook.SetSheetName(sheet.Workbook.GetSheetIndex(sheet), sheetName);
        }

        public static IRow GetRowWithPromise(this ISheet sheet, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }

            return row;
        }

        public static ICell GetCellWithPromise(this ISheet sheet, CellLocation cellLocation)
        {
            var row = sheet.GetRowWithPromise(cellLocation.RowIndex);
            var cell = row.GetCell(cellLocation.ColumnIndex);
            if (cell == null)
            {
                cell = row.CreateCell(cellLocation.ColumnIndex);
            }

            return cell;
        }

        public static ICell GetCell(this ISheet sheet, CellLocation cellLocation)
        {
            var row = sheet.GetRowWithPromise(cellLocation.RowIndex);

            return row.GetCell(cellLocation.ColumnIndex);
        }

        public static void SetCellValue(this ISheet sheet, CellLocation cellLocation, object value)
        {
            sheet.GetCellWithPromise(cellLocation).SetValue(value);
        }

        public static void RemoveCell(this ISheet sheet, CellLocation location)
        {
            var row = sheet.GetRow(location.RowIndex);
            if (row != null)
            {
                var cell = row.GetCell(location.ColumnIndex);
                if (cell != null)
                {
                    row.RemoveCell(cell);
                }
            }
        }

        /// <summary>
        /// 根据偏移量复制范围
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceRange"></param>
        /// <param name="offset"></param>
        public static void CopyRangeByOffset(this ISheet sheet, CellRange sourceRange, CellOffset offset)
        {
            var regionInfoList = sheet.GetMergedRegionInfos(sourceRange);
            sheet.RemoveMergedRegions(sourceRange);
            for (var rowIndex = sourceRange.EndLocation.RowIndex; rowIndex >= sourceRange.StartLocation.RowIndex; rowIndex--)
            {
                var targetRowIndex = rowIndex + offset.OffsetY;

                var sourceRow = sheet.GetRow(rowIndex);

                for (var columnIndex = sourceRange.EndLocation.ColumnIndex;
                    columnIndex >= sourceRange.StartLocation.ColumnIndex;
                    columnIndex--)
                {
                    var targetColumnIndex = columnIndex + offset.OffsetX;
                    var sourceCell = sourceRow?.GetCell(columnIndex);
                    sheet.CopyCell(sourceCell, new CellLocation(targetColumnIndex,targetRowIndex));
                }
            }

            foreach (MergedRegionInfo regionInfo in regionInfoList)
            {
                regionInfo.Range += offset;
                sheet.AddMergedRegion(regionInfo);
            }
        }

        public static void CopyCell(this ISheet sheet, ICell sourceCell, CellLocation targetLocation)
        {
            if (sourceCell == null)
            {
                sheet.RemoveCell(targetLocation);
                return;
            }

            var targetCell = sheet.GetCell(targetLocation);
            if (targetCell == null)
            {
                targetCell = sheet.GetCellWithPromise(targetLocation);
                if (targetCell.ColumnIndex != sourceCell.ColumnIndex)
                {
                    sheet.SetColumnWidth(targetCell.ColumnIndex, sheet.GetColumnWidth(sourceCell.ColumnIndex));
                }
            }

            if (sourceCell.CellStyle != null)
            {
                targetCell.CellStyle = sourceCell.CellStyle;
            }
            if (sourceCell.CellComment != null)
            {
                targetCell.CellComment = sourceCell.CellComment;
            }
            if (sourceCell.Hyperlink != null)
            {
                targetCell.Hyperlink = sourceCell.Hyperlink;
            }
            targetCell.SetCellType(sourceCell.CellType);

            switch (sourceCell.CellType)
            {
                case CellType.Numeric:
                    targetCell.SetCellValue(sourceCell.NumericCellValue);
                    break;
                case CellType.String:
                    targetCell.SetCellValue(sourceCell.RichStringCellValue);
                    break;
                case CellType.Formula:
                    targetCell.SetCellFormula(sourceCell.CellFormula);
                    break;
                case CellType.Blank:
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                    break;
                case CellType.Boolean:
                    targetCell.SetCellValue(sourceCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    targetCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                    break;
            }
        }

        public static void MoveRangeByOffset(this ISheet sheet, CellRange range, CellOffset offset)
        {
            sheet.CopyRangeByOffset(range, offset);
            var endRowIndex = range.StartLocation.RowIndex + offset.OffsetY;
            var endColumnIndex = range.StartLocation.ColumnIndex + offset.OffsetX;
            for (var rowIndex = range.StartLocation.RowIndex; rowIndex < endRowIndex; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }

                for (var columnIndex = range.StartLocation.ColumnIndex;
                    columnIndex <= range.EndLocation.ColumnIndex;
                    columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    if (cell == null)
                    {
                        continue;
                    }

                    row.RemoveCell(cell);
                }
            }

            for (var columnIndex = range.StartLocation.ColumnIndex; columnIndex < endColumnIndex; columnIndex++)
            {
                for (var rowIndex = range.StartLocation.RowIndex; rowIndex <= range.EndLocation.RowIndex; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null)
                    {
                        continue;
                    }

                    var cell = row.GetCell(columnIndex);
                    if (cell == null)
                    {
                        continue;
                    }

                    row.RemoveCell(cell);
                }
            }
        }

        public static void MoveRoundOfRangeByOffset(this ISheet sheet, CellRange range, CellOffset offset)
        {
            if (offset.OffsetX > 0)
            {
                var rightLastColumnIndex = sheet.GetLastColumnIndexForRoundOfRange(range);
                if (rightLastColumnIndex > range.EndLocation.ColumnIndex)
                {
                    sheet.MoveRangeByOffset(
                        new CellRange(range.EndLocation.ColumnIndex + 1,
                            range.StartLocation.RowIndex,
                            rightLastColumnIndex,
                            range.EndLocation.RowIndex), new CellOffset(offset.OffsetX));
                }
            }
            if (offset.OffsetY > 0)
            {
                var bottomLastRowIndex = sheet.LastRowNum;
                if (bottomLastRowIndex > range.EndLocation.RowIndex)
                {
                    sheet.MoveRangeByOffset(
                        new CellRange(range.StartLocation.ColumnIndex,
                            range.EndLocation.RowIndex + 1,
                            range.EndLocation.ColumnIndex + offset.OffsetX,
                            bottomLastRowIndex), offset);
                }
            }
        }

        public static int GetLastColumnIndexForRoundOfRange(this ISheet sheet, CellRange range)
        {
            var columnIndex = -1;
            for (var rowIndex = range.StartLocation.RowIndex; rowIndex < range.EndLocation.RowIndex; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row.LastCellNum > columnIndex)
                {
                    columnIndex = row.LastCellNum;
                }
            }

            return columnIndex;
        }

        /// <summary>
        /// 获取sheet中指定区域包含合并区域的信息列表
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static IList<MergedRegionInfo> GetMergedRegionInfos(this ISheet sheet, CellRange range)
        {
            var regionInfoList = new List<MergedRegionInfo>();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var mergedRegion = sheet.GetMergedRegion(i);
                var regionRange = new CellRange(mergedRegion.FirstColumn, mergedRegion.FirstRow, mergedRegion.LastColumn,
                    mergedRegion.LastRow);
                if (range.Include(regionRange))
                {
                    regionInfoList.Add(new MergedRegionInfo(i, regionRange));
                }
            }

            return regionInfoList;
        }

        /// <summary>
        /// 删除指定区域内的合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="range"></param>
        public static void RemoveMergedRegions(this ISheet sheet, CellRange range)
        {
            IList<MergedRegionInfo> regionInfoList;
            do
            {
                regionInfoList = sheet.GetMergedRegionInfos(range);
                var regionIndexs = regionInfoList.Select(r => r.Index).OrderByDescending(i => i);
                foreach (var ri in regionIndexs)
                {
                    sheet.RemoveMergedRegion(ri);
                }
            } while (regionInfoList.Count > 0);
        }

        /// <summary>
        /// 添加合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="regionInfo"></param>
        public static void AddMergedRegion(this ISheet sheet, MergedRegionInfo regionInfo)
        {
            var region = new CellRangeAddress(
                regionInfo.Range.StartLocation.RowIndex,
                regionInfo.Range.EndLocation.RowIndex,
                regionInfo.Range.StartLocation.ColumnIndex,
                regionInfo.Range.EndLocation.ColumnIndex);

            sheet.AddMergedRegion(region);
        }

        /// <summary>
        /// 添加多个合并区域
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="regionInfos"></param>
        public static void AddMergedRegions(this ISheet sheet, IEnumerable<MergedRegionInfo> regionInfos)
        {
            foreach (var regionInfo in regionInfos)
            {
                sheet.AddMergedRegion(regionInfo);
            }
        }
    }
}
