using ShapeCrawler.Shared;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table column.
/// </summary>
public interface IColumn
{
    /// <summary>
    ///     Gets or sets width in pixels.
    /// </summary>
    int Width { get; set; }
}

internal sealed class Column : IColumn
{
    internal Column(A.GridColumn aGridColumn)
    {
        this.AGridColumn = aGridColumn;
    }

    public int Width
    {
        get => this.GetWidth();
        set => this.SetWidth(value);
    }

    internal A.GridColumn AGridColumn { get; init; }

    private int GetWidth()
    {
        return new Emus(this.AGridColumn.Width!.Value).AsHorizontalPixels();
    }
    
    private void SetWidth(int pixels)
    {
        this.AGridColumn.Width!.Value = UnitConverter.HorizontalPixelToEmu(pixels);
    }
}