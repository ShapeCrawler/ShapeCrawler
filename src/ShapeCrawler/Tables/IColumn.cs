using System.Linq;
using ShapeCrawler.Shared;
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
    int Width { get; set; }
    
    /// <summary>
    ///     Creates a duplicate of the current column.
    /// </summary>
    void Duplicate();
}

internal sealed class Column : IColumn
{
    private readonly int index;

    internal Column(A.GridColumn aGridColumn, int index)
    {
        this.AGridColumn = aGridColumn;
        this.index = index;
    }

    public int Width
    {
        get => this.GetWidth();
        set => this.SetWidth(value);
    }

    internal A.GridColumn AGridColumn { get; }

    public void Duplicate()
    {
        var tableGrid = this.AGridColumn.Parent!;

        var existingColumns = tableGrid.Elements<A.GridColumn>().ToList();
        
        var totalWidth = existingColumns.Sum(col => col.Width!.Value);
        var newColumnWidth = totalWidth / (existingColumns.Count + 1);

        foreach (var col in existingColumns)
        {
            col.Width = newColumnWidth;
        }
        
        var newGridColumn = new A.GridColumn
        {
            Width = newColumnWidth
        };
        
        tableGrid.Append(newGridColumn);
        
        var table = tableGrid.Parent as A.Table;
        foreach(A.TableRow tr in table!.Elements<A.TableRow>())
        {
            var cells = tr.Elements<A.TableCell>().ToList();
            var cloneCell = cells[this.index].Clone();
            tr.InsertAfter((A.TableCell)cloneCell, cells.Last());
        }
    }

    private int GetWidth()
    {
        return new Emus(this.AGridColumn.Width!.Value).AsHorizontalPixels();
    }
    
    private void SetWidth(int pixels)
    {
        this.AGridColumn.Width!.Value = UnitConverter.HorizontalPixelToEmu(pixels);
    }
}