// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table cell.
/// </summary>
public interface ITableCell
{
    /// <summary>
    ///     Gets text box.
    /// </summary>
    ITextFrame TextFrame { get; }

    /// <summary>
    ///     Gets a value indicating whether cell belongs to merged cell.
    /// </summary>
    bool IsMergedCell { get; }
}