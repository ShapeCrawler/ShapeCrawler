namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft for configuring a solid fill.
/// </summary>
public sealed class DraftSolidFill
{
    internal string? HexColor { get; private set; }

    internal int? TransparencyPercent { get; private set; }

    /// <summary>
    ///     Sets fill color.
    /// </summary>
    /// <param name="hexColor">Hex color string (e.g., "0000FF").</param>
    public DraftSolidFill Color(string hexColor)
    {
        this.HexColor = hexColor;
        return this;
    }

    /// <summary>
    ///     Sets fill transparency in percents. Range is 0 (opaque) to 100 (fully transparent).
    /// </summary>
    public DraftSolidFill Transparency(int transparencyPercent)
    {
        this.TransparencyPercent = transparencyPercent;
        return this;
    }
}