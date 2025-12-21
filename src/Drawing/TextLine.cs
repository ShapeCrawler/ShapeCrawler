namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents a text line.
/// </summary>
internal record TextLine(
    PixelTextPortion[] Runs,
    TextHorizontalAlignment HorizontalAlignment,
    float ParaLeftMargin,
    float Width,
    float Height,
    float BaselineOffset);
