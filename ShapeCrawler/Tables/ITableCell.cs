using ShapeCrawler.AutoShapes;
// ReSharper disable CheckNamespace

namespace ShapeCrawler
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