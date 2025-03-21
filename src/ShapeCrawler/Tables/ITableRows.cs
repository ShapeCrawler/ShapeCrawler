using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

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
    ///     Gets row at the specified index.
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
    ///     Adds a new row at the end of table.
    /// </summary>
    void Add();
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

    IEnumerator<ITableRow> IEnumerable<ITableRow>.GetEnumerator() => this.Rows().GetEnumerator();
    
    IEnumerator IEnumerable.GetEnumerator() => this.Rows().GetEnumerator();

    private List<TableRow> Rows() => [.. this.aTable.Elements<A.TableRow>().Select((aTableRow, index) => new TableRow(aTableRow, index))];
}