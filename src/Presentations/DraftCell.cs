namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft cell builder.
/// </summary>
public sealed class DraftCell
{
    internal string? SolidColorHex { get; private set; }

    /// <summary>
    ///     Sets the solid color fill for the cell.
    /// </summary>
    public DraftCell SolidColor(string hex)
    {
        this.SolidColorHex = hex;
        return this;
    }
}
