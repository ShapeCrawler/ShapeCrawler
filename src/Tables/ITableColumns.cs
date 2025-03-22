using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

/// <summary>
///    Represents a table column collection.
/// </summary>
public interface ITableColumns : IEnumerable<IColumn>
{
    /// <summary>
    ///     Gets number of columns.
    /// </summary>
    int Count { get; }
    
    /// <summary>
    ///     Gets column at the specified index.
    /// </summary>
    IColumn this[int index] { get; }
    
    /// <summary>
    ///     Removes specified column from collection.
    /// </summary>
    void Remove(IColumn removing);
    
    /// <summary>
    ///     Removes column by index.
    /// </summary>
    void RemoveAt(int index);
    
    /// <summary>
    ///     Adds a new column at the end of table.
    /// </summary>
    void Add();
    
    /// <summary>
    ///     Inserts a new column after the specified column.
    /// </summary>
    /// <param name="columnNumber">The column number after which to add the new column.</param>
    void InsertAfter(int columnNumber);
}

internal sealed class TableColumns : ITableColumns
{
    private readonly DocumentFormat.OpenXml.Drawing.Table aTable;

    internal TableColumns(DocumentFormat.OpenXml.Presentation.GraphicFrame pGraphicFrame)
    {
        this.aTable = pGraphicFrame.GetFirstChild<DocumentFormat.OpenXml.Drawing.Graphic>() !.GraphicData!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Table>() !;
    }
    
    public int Count => this.Columns().Count;

    public IColumn this[int index] => this.Columns()[index];

    public void Remove(IColumn removing)
    {
        var internalCol = (Column)removing;
        internalCol.AGridColumn.Remove();
    }

    public void RemoveAt(int index)
    {
        var cols = this.Columns();
        if (index < 0 || index >= cols.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }
        
        var internalCol = cols[index];
        this.Remove(internalCol);
    }

    public void Add() => this.InsertAfter(this.Columns().Count);
    
    public void InsertAfter(int columnNumber)
    {
        var columnIndex = columnNumber - 1;
        var tableGrid = this.aTable.TableGrid!;
        var existingColumns = this.Columns().Select(x => x.AGridColumn).ToList();

        var gridColumn = Column.CreateNewColumn(tableGrid, existingColumns[columnIndex].Width!.Value);
        var targetColumn = existingColumns[columnIndex];
        tableGrid.InsertAfter(gridColumn, targetColumn);

        foreach (var aTableRow in this.aTable.Elements<A.TableRow>())
        {
            new SCATableRow(aTableRow).InsertNewCellAfter(columnNumber);
        }
    }
    
    IEnumerator<IColumn> IEnumerable<IColumn>.GetEnumerator() => this.Columns().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.Columns().GetEnumerator();

    private List<Column> Columns() => [.. this.aTable.TableGrid!.Elements<A.GridColumn>().Select((aGridColumn, index) => new Column(aGridColumn, index))];
}