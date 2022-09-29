using System.Collections.Generic;

namespace MyWpf.Controls.Ribbon
{
    internal static class CollectionExtensions
    {
        public static void AddRange<T>(this ICollection<T> target, IEnumerable<T> source)
        {
            foreach (var item in source)
            {
                target.Add(item);
            }
        }

        //public static void AddRange<T>(this ItemCollection target, IEnumerable<T> source)
        //{
        //    foreach (var item in source)
        //    {
        //        target.Add(item);
        //    }
        //}
    }
}
