using System.Collections.Generic;
using System.Linq;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents collection of text lines.
/// </summary>
internal sealed class TextLines(IList<TextLine> lines)
{
    /// <summary>
    ///     Renders text lines on the canvas.
    /// </summary>
    internal void Render(SKCanvas canvas, float x, float y, float availableWidth, float availableHeight, TextVerticalAlignment verticalAlignment)
    {
        var textBlockHeight = lines.Sum(l => l.Height);
        var verticalOffset = DrawingTextBox.GetVerticalOffset(verticalAlignment, availableHeight, textBlockHeight);
        var lineTop = y + verticalOffset;

        foreach (var line in lines)
        {
            var horizontalOffset = DrawingTextBox.GetHorizontalOffset(line.HorizontalAlignment, availableWidth - line.ParaLeftMargin, line.Width);
            var startX = x + line.ParaLeftMargin + horizontalOffset;

            RenderLine(canvas, line, startX, lineTop);
            lineTop += line.Height;
        }
    }

    private static void RenderLine(SKCanvas canvas, TextLine line, float startX, float lineTop)
    {
        var baselineY = lineTop + line.BaselineOffset;
        var currentX = startX;

        foreach (var run in line.Runs)
        {
            using var font = DrawingTextBox.CreateFont(run.Font);
            using var paint = DrawingTextBox.CreatePaint(run.Font);

            canvas.DrawText(run.Text, currentX, baselineY, SKTextAlign.Left, font, paint);
            currentX += run.Width;
        }
    }
}
