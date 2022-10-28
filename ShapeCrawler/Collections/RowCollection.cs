using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Collections;

/// <summary>
///     Represent a collect of table rows.
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
    void Remove(ITableRow row);

    /// <summary>
    ///     Removes table row by index.
    /// </summary>
    void RemoveAt(int index);
}

internal class RowCollection : IRowCollection
{
    private readonly List<SCTableRow> collectionItems;

    #region Constructors

    private RowCollection(List<SCTableRow> rowList)
    {
        this.collectionItems = rowList;
    }

    #endregion Constructors

    public int Count => this.collectionItems.Count;

    public ITableRow this[int index] => this.collectionItems[index];

    public void Remove(ITableRow removingRow)
    {
        var removingRowInternal = (SCTableRow)removingRow;
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

    IEnumerator<ITableRow> IEnumerable<ITableRow>.GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }

    internal static RowCollection Create(SlideTable table, P.GraphicFrame pGraphicFrame)
    {
        IEnumerable<A.TableRow> aTableRows = pGraphicFrame.GetATable().Elements<A.TableRow>();
        var rowList = new List<SCTableRow>(aTableRows.Count());
        int rowIndex = 0;
        rowList.AddRange(aTableRows.Select(aTblRow => new SCTableRow(table, aTblRow, rowIndex++)));

        return new RowCollection(rowList);
    }
}