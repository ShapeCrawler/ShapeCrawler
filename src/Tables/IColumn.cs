using System;
using System.Linq;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents a table column.
/// </summary>
public interface IColumn
{
    /// <summary>
    ///     Gets or sets width in pixels.
    /// </summary>
    decimal Width { get; set; }
    
    /// <summary>
    ///     Creates a duplicate of the current column at the end of the table.
    /// </summary>
    void Duplicate();
}

internal sealed class Column(A.GridColumn aGridColumn, int index): IColumn
{
    public decimal Width
    {
        get => new Emus(this.AGridColumn.Width!.Value).AsPoints();
        set => this.AGridColumn.Width!.Value = new Points(value).AsEmus();
    }

    internal A.GridColumn AGridColumn => aGridColumn;

    public void Duplicate()
    {
        var tableGrid = this.AGridColumn.Parent as A.TableGrid;
        
        var newGridColumn = CreateNewColumn(tableGrid!, this.AGridColumn.Width!.Value);
        
        tableGrid!.Append(newGridColumn);
        
        var table = tableGrid.Parent as A.Table;
        
        foreach(A.TableRow tr in table!.Elements<A.TableRow>())
        {
            var cells = tr.Elements<A.TableCell>().ToList();
            var cloneCell = cells[index].Clone();
            tr.InsertAfter((A.TableCell)cloneCell, cells[^1]);
        }
    }
    
    internal static A.GridColumn CreateNewColumn(A.TableGrid tableGrid, long width)
    {
        var totalWidth = tableGrid.Elements<A.GridColumn>().Sum(col => col.Width!.Value);
        var newTotalWidth = totalWidth + width;
        var ratio = (double)totalWidth / newTotalWidth;

        var newGridColumn = new A.GridColumn { Width = (int)Math.Round(width * ratio) };

        foreach (var col in tableGrid.Elements<A.GridColumn>())
        {
            col.Width = (int)Math.Round(col.Width!.Value * ratio);
        }

        return newGridColumn;
    }
}