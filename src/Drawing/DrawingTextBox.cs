using DocumentFormat.OpenXml;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents drawing text box.
/// </summary>
internal sealed class DrawingTextBox : TextBox
{
    internal DrawingTextBox(TextBoxMargins margins, OpenXmlElement textBody)
        : base(margins, textBody)
    {
    }

    internal void Render(SKCanvas canvas, decimal parentShapeX, decimal parentShapeY, decimal parentShapeWidth, decimal parentShapeHeight)
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            return;
        }

        var originX = (float)new Points(parentShapeX + LeftMargin).AsPixels();
        var originY = (float)new Points(parentShapeY + TopMargin).AsPixels();
        var availableWidth = GetAvailableWidth(parentShapeWidth);
        var availableHeight = GetAvailableHeight(parentShapeHeight);

        var wrap = this.TextWrapped && availableWidth > 0;
        new DrawingTextLines(this.Paragraphs, availableWidth, wrap).Render(canvas, originX, originY, availableHeight, VerticalAlignment);
    }
    
    private static decimal ClampToZero(decimal value)
    {
        return value < 0 ? 0 : value;
    }

    private float GetAvailableWidth(decimal parentShapeWidth)
    {
        var width = ClampToZero(parentShapeWidth - this.LeftMargin - this.RightMargin);
        return (float)new Points(width).AsPixels();
    }

    private float GetAvailableHeight(decimal parentShapeHeight)
    {
        var height = ClampToZero(parentShapeHeight - this.TopMargin - this.BottomMargin);
        return (float)new Points(height).AsPixels();
    }
}
