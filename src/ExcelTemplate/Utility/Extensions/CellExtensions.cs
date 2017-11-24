using System;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Utility.Extensions
{
    public static class CellExtensions
    {
        public static void SetValue(this ICell cell, object value)
        {
            if (null == cell)
            {
                return;
            }
            if (null == value)
            {
                cell.SetCellValue(string.Empty);
            }
            else
            {
                TypeCode valueTypeCode = Type.GetTypeCode(value.GetType());
                switch (valueTypeCode)
                {
                    case TypeCode.String:
                        if (value.ToString().Contains("\n"))
                        {
                            cell.CellStyle.WrapText = true;
                        }
                        cell.SetCellValue(System.Convert.ToString(value));
                        break;

                    case TypeCode.DateTime:
                        cell.SetCellValue(System.Convert.ToDateTime(value));
                        break;

                    case TypeCode.Boolean:
                        cell.SetCellValue(System.Convert.ToBoolean(value));
                        break;

                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                    case TypeCode.Byte:
                    case TypeCode.Single:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                    case TypeCode.UInt16:
                    case TypeCode.UInt32:
                    case TypeCode.UInt64:
                        cell.SetCellValue(System.Convert.ToDouble(value));
                        break;

                    default:
                        cell.SetCellValue(string.Empty);
                        break;
                }
            }
        }
    }
}
