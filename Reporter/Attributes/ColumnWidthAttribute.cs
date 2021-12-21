namespace Reporter.Attributes
{
    /// <summary>
    /// 欄位寬度
    /// </summary>
    [AttributeUsage(AttributeTargets.All, AllowMultiple = false)]
    public class ColumnWidthAttribute : Attribute
    {
        /// <summary>
        /// 建構子
        /// </summary>
        public ColumnWidthAttribute() { }
        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="columnWidth"></param>
        public ColumnWidthAttribute(int columnWidth)
        {
            ColumnWidth = columnWidth;
        }
        /// <summary>
        /// 報表欄寬
        /// </summary>
        public int ColumnWidth { get; set; }
    }
}
