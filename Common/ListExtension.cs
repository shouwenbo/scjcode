namespace Common
{
    public static class ListExtension
    {
        public static bool CheckAsc<T, TKey>(this List<T> list, Func<T, TKey> keySelector) where TKey : IComparable<TKey>
        {
            if (list.OrderBy(keySelector).SequenceEqual(list))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool CheckDesc<T, TKey>(this List<T> list, Func<T, TKey> keySelector) where TKey : IComparable<TKey>
        {
            if (list.OrderByDescending(keySelector).SequenceEqual(list))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}