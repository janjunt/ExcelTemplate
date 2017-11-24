using System.Collections.Generic;

namespace ExcelTemplate.Utility.Extensions
{
    /// <summary>
    /// 集合扩展
    /// </summary>
    public static class CollectionExtensions
    {
        /// <summary>
        /// 把集合添加到原集合中
        /// </summary>
        /// <typeparam name="T">元素项类型</typeparam>
        /// <param name="source">原集合</param>
        /// <param name="collection">要添加的集合</param>
        public static void AddRange<T>(this ICollection<T> source, IEnumerable<T> collection)
        {
            if (source == null || collection == null)
            {
                return;
            }

            foreach (var item in collection)
            {
                source.Add(item);
            }
        }
    }
}
