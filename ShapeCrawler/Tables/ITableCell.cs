using ShapeCrawler.AutoShapes;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a table cell.
    /// </summary>
    public interface ITableCell
    {
        /// <summary>
        ///     Gets text box.
        /// </summary>
        ITextBox TextBox { get; }

        /// <summary>
        ///     Gets a value indicating whether cell belongs to merged cell.
        /// </summary>
        bool IsMergedCell { get; }
    }
}