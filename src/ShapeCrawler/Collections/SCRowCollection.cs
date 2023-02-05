using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table row collection.
/// </summary>
public interface IRowCollection : IEnumerable<IRow>
{
    /// <summary>
    ///     Gets number of rows.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets row at the specified index.
    /// </summary>
    IRow this[int index] { get; }

    /// <summary>
    ///     Removes specified row from collection.
    /// </summary>
    void Remove(IRow row);

    /// <summary>
    ///     Removes table row by index.
    /// </summary>
    void RemoveAt(int index);

    /// <summary>
    ///     Adds a new row at the end of table.
    /// </summary>
    IRow Add();
}

internal sealed class SCRowCollection : IRowCollection
{
    private readonly List<SCRow> collectionItems;
    private readonly SCTable parentTable;
    private readonly A.Table aTable;

    private SCRowCollection(List<SCRow> rowList, SCTable parentTable, A.Table aTable)
    {
        this.collectionItems = rowList;
        this.parentTable = parentTable;
        this.aTable = aTable;
    }

    public int Count => this.collectionItems.Count;

    public IRow this[int index] => this.collectionItems[index];

    public void Remove(IRow removingRow)
    {
        var removingRowInternal = (SCRow)removingRow;
        removingRowInternal.ATableRow.Remove();
        this.collectionItems.Remove(removingRowInternal);
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= this.collectionItems.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var innerRow = this.collectionItems[index];
        this.Remove(innerRow);
    }

    public IRow Add()
    {
        var columnsCount = this.collectionItems[0].Cells.Count;
        var aTableRow = this.aTable.AddRow(columnsCount);
        var tableRow = new SCRow(this.parentTable, aTableRow, this.collectionItems.Count);
        this.collectionItems.Add(tableRow);

        return tableRow;
    }

    IEnumerator<IRow> IEnumerable<IRow>.GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }

    internal static SCRowCollection Create(SCTable table, P.GraphicFrame pGraphicFrame)
    {
        var aTable = pGraphicFrame.GetATable();
        var aTableRows = aTable.Elements<A.TableRow>();
        var rowList = new List<SCRow>(aTableRows.Count());
        var rowIndex = 0;
        rowList.AddRange(aTableRows.Select(aTblRow => new SCRow(table, aTblRow, rowIndex++)));

        return new SCRowCollection(rowList, table, aTable);
    }
}