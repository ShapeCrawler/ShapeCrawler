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
    ITableRow Add();
}

internal sealed class SCRowCollection : IRowCollection
{
    private readonly List<SCTableRow> collectionItems;
    private readonly SlideTable parentTable;
    private readonly A.Table aTable;
    private readonly TypedOpenXmlPart slideTypedOpenXmlPart;
    private readonly List<ImagePart> imageParts;

    private SCRowCollection(List<SCTableRow> rowList, SlideTable parentTable, A.Table aTable, TypedOpenXmlPart slideTypedOpenXmlPart, List<ImagePart> imageParts)
    {
        this.collectionItems = rowList;
        this.parentTable = parentTable;
        this.aTable = aTable;
        this.slideTypedOpenXmlPart = slideTypedOpenXmlPart;
        this.imageParts = imageParts;
    }

    public int Count => this.collectionItems.Count;

    public ITableRow this[int index] => this.collectionItems[index];

    public void Remove(ITableRow removingTableRow)
    {
        var removingRowInternal = (SCTableRow)removingTableRow;
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

    public ITableRow Add()
    {
        var columnsCount = this.collectionItems[0].Cells.Count;
        var aTableRow = this.aTable.AddRow(columnsCount);
        var tableRow = new SCTableRow(this.parentTable, aTableRow, this.collectionItems.Count, this.slideTypedOpenXmlPart, this.imageParts);
        this.collectionItems.Add(tableRow);

        return tableRow;
    }

    IEnumerator<ITableRow> IEnumerable<ITableRow>.GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }

    internal static SCRowCollection Create(
        SlideTable table, 
        P.GraphicFrame pGraphicFrame, 
        TypedOpenXmlPart slideTypedOpenXmlPart, 
        List<ImagePart> imageParts)
    {
        var aTable = pGraphicFrame.GetATable();
        var aTableRows = aTable.Elements<A.TableRow>();
        var rowList = new List<SCTableRow>(aTableRows.Count());
        var rowIndex = 0;
        rowList.AddRange(aTableRows.Select(aTblRow => new SCTableRow(table, aTblRow, rowIndex++, slideTypedOpenXmlPart, imageParts)));

        return new SCRowCollection(rowList, table, aTable, slideTypedOpenXmlPart, imageParts);
    }
}