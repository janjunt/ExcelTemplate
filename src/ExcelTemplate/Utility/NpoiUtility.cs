using NPOI.SS.UserModel;
using System.IO;

namespace ExcelTemplate.Utility
{
    /// <summary>
    /// Npoi操作单元
    /// </summary>
    public static class NpoiUtility
    {
        /// <summary>
        /// 加载Excel文件,获取IWorkbook对象
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>IWorkbook对象</returns>
        public static IWorkbook LoadWorkbook(string filePath)
        {
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return WorkbookFactory.Create(fileStream);
            }
        }
    }
}
