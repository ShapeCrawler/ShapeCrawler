using ShapeCrawler.Statics;
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

internal class SCColumn : IColumn
{
    internal SCColumn(A.GridColumn aGridColumn)
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
        return UnitConverter.HorizontalEmuToPixel(this.AGridColumn.Width!.Value);
    }
    
    private void SetWidth(int pixels)
    {
        this.AGridColumn.Width!.Value = UnitConverter.HorizontalPixelToEmu(pixels);
    }
}