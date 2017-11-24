namespace ExcelTemplate.Utility.Models
{
    /// <summary>
    /// 单元格范围
    /// </summary>
    public class CellRange
    {
        public CellRange(CellLocation startLocation, CellLocation endLocation)
        {
            StartLocation = startLocation;
            EndLocation = endLocation;
        }

        public CellRange(int startColumnIndex, int startRowIndex, int endColumnIndex, int endRowIndex)
            : this(new CellLocation(startColumnIndex, startRowIndex), new CellLocation(endColumnIndex, endRowIndex))
        {
        }

        /// <summary>
        /// 起始单元格
        /// </summary>
        public CellLocation StartLocation { get; set; }
        /// <summary>
        /// 终止单元格
        /// </summary>
        public CellLocation EndLocation { get; set; }
        /// <summary>
        /// 列数
        /// </summary>
        public int ColumnNumber {
            get { return EndLocation.ColumnIndex - StartLocation.ColumnIndex + 1; }
        }
        /// <summary>
        /// 行数
        /// </summary>
        public int RowNumber
        {
            get { return EndLocation.RowIndex - StartLocation.RowIndex + 1; }
        }
        public CellRange Copy()
        {
            return new CellRange(
                StartLocation.ColumnIndex, 
                StartLocation.RowIndex, 
                EndLocation.ColumnIndex,
                EndLocation.RowIndex);
        }


        public static CellRange operator +(CellRange range, CellOffset offset)
        {
            return new CellRange(range.StartLocation + offset, range.EndLocation + offset);
        }
    }

    public static class CellRangeExtensions
    {
        public static bool HasCross(this CellRange source, CellRange other)
        {
            return !(source.StartLocation.ColumnIndex > other.EndLocation.ColumnIndex ||
                     source.EndLocation.ColumnIndex < other.StartLocation.ColumnIndex ||
                     source.StartLocation.RowIndex > other.EndLocation.RowIndex ||
                     source.EndLocation.RowIndex < other.StartLocation.RowIndex);
        }

        public static bool Include(this CellRange range, CellLocation location)
        {
            return location.ColumnIndex >= range.StartLocation.ColumnIndex &&
                   location.ColumnIndex <= range.EndLocation.ColumnIndex &&
                   location.RowIndex >= range.StartLocation.RowIndex &&
                   location.RowIndex <= range.EndLocation.RowIndex;
        }

        public static bool Include(this CellRange range, CellRange otherRange)
        {
            return otherRange.StartLocation.ColumnIndex >= range.StartLocation.ColumnIndex &&
                   otherRange.EndLocation.ColumnIndex <= range.EndLocation.ColumnIndex &&
                   otherRange.StartLocation.RowIndex >= range.StartLocation.RowIndex &&
                   otherRange.EndLocation.RowIndex <= range.EndLocation.RowIndex;
        }

        public static bool Include(this CellRange range, int startColumnIndex, int startRowIndex, int endColumnIndex, int endRowIndex)
        {
            return range.Include(new CellRange(startColumnIndex, startRowIndex, endColumnIndex, endRowIndex));
        }
    }
}
