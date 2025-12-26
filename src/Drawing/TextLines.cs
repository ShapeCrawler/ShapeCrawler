using System.Collections.Generic;
using System.Linq;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents text lines.
/// </summary>
/// <param name="lines">The text lines.</param>
/// <param name="availableWidth">The maximum available width for each rendered line.</param>
internal sealed class TextLines(IReadOnlyList<TextLine> lines, float availableWidth)
{
    private readonly IReadOnlyList<TextLine> textLines = lines;
    private readonly float availableWidth = availableWidth;

    /// <summary>
    ///     Renders the text lines.
    /// </summary>
    /// <param name="canvas">The canvas where the text is rendered.</param>
    /// <param name="x">The x coordinate of the text block.</param>
    /// <param name="y">The y coordinate of the text block.</param>
    /// <param name="availableHeight">The maximum available height.</param>
    /// <param name="verticalAlignment">The vertical alignment.</param>
    internal void Render(SKCanvas canvas, float x, float y, float availableHeight, TextVerticalAlignment verticalAlignment)
    {
        var textBlockHeight = this.textLines.Sum(l => l.Height);
        var verticalOffset = GetVerticalOffset(verticalAlignment, availableHeight, textBlockHeight);
        var lineTop = y + verticalOffset;

        foreach (var textLine in this.textLines)
        {
            var horizontalOffset = GetHorizontalOffset(
                textLine.HorizontalAlignment,
                this.availableWidth - textLine.ParaLeftMargin,
                textLine.Width);
            var startX = x + textLine.ParaLeftMargin + horizontalOffset;

            textLine.Render(canvas, startX, lineTop);
            lineTop += textLine.Height;
        }
    }

    private static float GetVerticalOffset(TextVerticalAlignment alignment, float availableHeight, float contentHeight)
    {
        if (availableHeight <= 0)
        {
            return 0;
        }

        var freeSpace = availableHeight - contentHeight;
        return alignment switch
        {
            TextVerticalAlignment.Top => 0,
            TextVerticalAlignment.Middle => freeSpace / 2,
            TextVerticalAlignment.Bottom => freeSpace,
            _ => 0
        };
    }

    private static float GetHorizontalOffset(TextHorizontalAlignment alignment, float availableWidth, float lineWidth)
    {
        if (availableWidth <= 0)
        {
            return 0;
        }

        var freeSpace = availableWidth - lineWidth;
        return alignment switch
        {
            TextHorizontalAlignment.Left => 0,
            TextHorizontalAlignment.Center => freeSpace / 2,
            TextHorizontalAlignment.Right => freeSpace,
            _ => 0 // Treat Justify and unknown values as Left for MVP.
        };
    }
}