namespace ExcelTemplate.Utility.Models
{
    public class MergedRegionInfo
    {
        public MergedRegionInfo(int index, CellRange range)
        {
            Index = index;
            Range = range;
        }

        public MergedRegionInfo(int index, CellLocation startLocation, CellLocation endLocation)
            : this(index, new CellRange(startLocation, endLocation))
        {
        }

        public MergedRegionInfo(int index, int startColumnIndex, int startRowIndex, int endColumnIndex, int endRowIndex)
            : this(index, new CellRange(startColumnIndex, startRowIndex, endColumnIndex, endRowIndex))
        {
        }

        public int Index { get; set; }

        public CellRange Range { get; set; }
    }
}
