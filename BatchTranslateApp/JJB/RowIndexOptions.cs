namespace BatchTranslateApp
{
    public class RowIndexOptions
    {
        public int StartColumn { get; set; }
        public List<RowIndexItem> List { get; set; }
    }

    public class RowIndexItem
    {
        /// <summary>
        /// 行名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 行索引
        /// </summary>
        public int Index { get; set; }
    }
}
