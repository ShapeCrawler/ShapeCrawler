using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Tables
{
    public interface ITableCell
    {
        /// <summary>
        ///     Gets text box.
        /// </summary>
        ITextBox TextBox { get; }

        bool IsMergedCell { get; }
    }
}