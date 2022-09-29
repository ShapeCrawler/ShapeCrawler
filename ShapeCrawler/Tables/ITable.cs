using System.Collections.Generic;
using ShapeCrawler.Collections;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table on a slide.
/// </summary>
public interface ITable : IShape
{
    /// <summary>
    ///     Gets table columns.
    /// </summary>
    IReadOnlyList<SCColumn> Columns { get; }

    /// <summary>
    ///     Gets table rows.
    /// </summary>
    IRowCollection Rows { get; }

    /// <summary>
    ///     Gets cell by row and column indexes.
    /// </summary>
    ITableCell this[int rowIndex, int columnIndex] { get; }

    /// <summary>
    ///     Merge neighbor cells.
    /// </summary>
    void MergeCells(ITableCell cell1, ITableCell cell2);

    /// <summary>
    ///     Removes row at the specified index.
    /// </summary>
    /// <param name="index">The index of the row that should be removed.</param>
    void RemoveRowAt(int index); // TODO: move to row collection

#if DEBUG
    
    /// <summary>
    ///     Adds specified row at the bottom of the current table.
    /// </summary>
    /// <param name="aTableRow">Row that will be added to the table.</param>
    /// <returns>A reference to the recently added row.</returns>
    ITableRow AppendRow(A.TableRow aTableRow); // TODO: move to row collection
    
#endif
}