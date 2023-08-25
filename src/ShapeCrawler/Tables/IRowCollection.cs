using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table row collection.
/// </summary>
public interface IRowCollection : IEnumerable<ITableRow>
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
    void Remove(ITableRow tableRow);

    /// <summary>
    ///     Removes table row by index.
    /// </summary>
    void RemoveAt(int index);

    /// <summary>
    ///     Adds a new row at the end of table.
    /// </summary>
    void Add();
}

internal sealed class SlideTableRows : IRowCollection
{
    private readonly SlidePart sdkSlidePart;
    private readonly List<SlideTableRow> rows;
    private readonly A.Table aTable;
    private readonly P.GraphicFrame pGraphicFrame;

    internal SlideTableRows (SlidePart sdkSlidePart, P.GraphicFrame pGraphicFrame)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pGraphicFrame = pGraphicFrame;
        var aTable = pGraphicFrame.GetATable();
        var aTableRows = aTable.Elements<A.TableRow>();
        var rowList = new List<SlideTableRow>(aTableRows.Count());
        var rowIndex = 0;
        rowList.AddRange(aTableRows.Select(aTblRow => new SlideTableRow(this.sdkSlidePart, aTblRow, rowIndex++)));

        this.rows = rowList;
        this.aTable = aTable;
    }

    public int Count => this.rows.Count;

    public ITableRow this[int index] => this.rows[index];

    public void Remove(ITableRow removingTableRow)
    {
        var removingRowInternal = (SlideTableRow)removingTableRow;
        removingRowInternal.ATableRow.Remove();
        this.rows.Remove(removingRowInternal);
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= this.rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var innerRow = this.rows[index];
        this.Remove(innerRow);
    }

    public void Add()
    {
        var columnsCount = this.rows[0].Cells.Count;
        var aTableRow = this.aTable.AddRow(columnsCount);
        var tableRow = new SlideTableRow(this.sdkSlidePart, aTableRow, this.rows.Count);
        this.rows.Add(tableRow);
    }

    IEnumerator<ITableRow> IEnumerable<ITableRow>.GetEnumerator()
    {
        return this.rows.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.rows.GetEnumerator();
    }
}