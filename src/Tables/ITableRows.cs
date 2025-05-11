using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a table row collection.
/// </summary>
public interface ITableRows : IEnumerable<ITableRow>
{
    /// <summary>
    ///     Gets number of rows.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets row by index.
    /// </summary>
    ITableRow this[int index] { get; }

    /// <summary>
    ///     Removes specified row from collection.
    /// </summary>
    void Remove(ITableRow removing);

    /// <summary>
    ///     Removes table row by index.
    /// </summary>
    void RemoveAt(int index);

    /// <summary>
    ///     Adds a new row at the end of the table.
    /// </summary>
    void Add();

#if DEBUG
    /// <summary>
    ///     Adds a new row at the specified index.
    /// </summary>
    void Add(int index);

    /// <summary>
    ///     Adds a new row at the specified index.
    /// </summary>
    /// <param name="index">Index where the new row will be added.</param>
    /// <param name="templateRowIndex">Row index used as a format template for the new row.</param>
    void Add(int index, int templateRowIndex);
#endif
}

internal sealed class TableRows : ITableRows
{
    private readonly A.Table aTable;

    internal TableRows(P.GraphicFrame pGraphicFrame)
    {
        this.aTable = pGraphicFrame.GetFirstChild<A.Graphic>() !.GraphicData!.GetFirstChild<A.Table>() !;
    }

    public int Count => this.Rows().Count;

    public ITableRow this[int index] => this.Rows()[index];

    public void Remove(ITableRow removing)
    {
        var removingRowInternal = (TableRow)removing;
        removingRowInternal.ATableRow.Remove();
    }

    public void RemoveAt(int index)
    {
        var rows = this.Rows();
        if (index < 0 || index >= rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var innerRow = rows[index];
        this.Remove(innerRow);
    }

    public void Add()
    {
        var columnsCount = this.Rows()[0].Cells.Count;
        this.aTable.AddRow(columnsCount);
    }

#if DEBUG
    public void Add(int index)
    {
        var rows = this.Rows();
        if (index < 0 || index > rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var columnsCount = rows.Count > 0 ? rows[0].Cells.Count : 0;
        if (columnsCount == 0)
        {
            throw new InvalidOperationException("Cannot add a row to an empty table.");
        }

        var aTableRow = new A.TableRow { Height = Constants.DefaultRowHeightEmu };
        for (var i = 0; i < columnsCount; i++)
        {
            new SCATableRow(aTableRow).AddNewCell();
        }
        
        // Get the element before which we want to insert the new row
        var aTableRows = this.aTable.Elements<A.TableRow>().ToList();
        if (index == aTableRows.Count)
        {
            // Add at the end
            this.aTable.Append(aTableRow);
        }
        else
        {
            // Insert before the row at the specified index
            this.aTable.InsertBefore(aTableRow, aTableRows[index]);
        }
    }

    public void Add(int index, int templateRowIndex)
    {
        var rows = this.Rows();
        if (index < 0 || index > rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        if (templateRowIndex < 0 || templateRowIndex >= rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(templateRowIndex));
        }

        // Clone the template row
        var templateRow = (TableRow)rows[templateRowIndex];
        var templateARow = templateRow.ATableRow;
        var newARow = (A.TableRow)templateARow.CloneNode(true);
        
        // Get the element before which we want to insert the new row
        var aTableRows = this.aTable.Elements<A.TableRow>().ToList();
        if (index == aTableRows.Count)
        {
            // Add at the end
            this.aTable.Append(newARow);
        }
        else
        {
            // Insert before the row at the specified index
            this.aTable.InsertBefore(newARow, aTableRows[index]);
        }
    }
#endif

    IEnumerator<ITableRow> IEnumerable<ITableRow>.GetEnumerator() => this.Rows().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.Rows().GetEnumerator();

    private List<TableRow> Rows() =>
    [
        .. this.aTable.Elements<A.TableRow>().Select((aTableRow, index) => new TableRow(aTableRow, index))
    ];
}