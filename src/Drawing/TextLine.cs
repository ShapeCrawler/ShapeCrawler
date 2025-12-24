using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents a text line.
/// </summary>
internal sealed class TextLine(
    PixelTextPortion[] runs,
    TextHorizontalAlignment horizontalAlignment,
    float paraLeftMargin,
    float width,
    float height,
    float baselineOffset)
{
    internal TextHorizontalAlignment HorizontalAlignment => horizontalAlignment;

    internal float ParaLeftMargin => paraLeftMargin;

    internal float Width => width;

    internal float Height => height;

    private float BaselineOffset => baselineOffset;

    private PixelTextPortion[] Runs => runs;

    internal void Render(SKCanvas canvas, float x, float y)
    {
        var baselineY = y + this.BaselineOffset;
        var currentX = x;

        foreach (var run in this.Runs)
        {
            var drawingFont = new DrawingFont(run.Font);
            using var font = drawingFont.AsSkFont();
            using var paint = drawingFont.CreatePaint();

            canvas.DrawText(run.Text, currentX, baselineY, SKTextAlign.Left, font, paint);
            currentX += run.Width;
        }
    }
}