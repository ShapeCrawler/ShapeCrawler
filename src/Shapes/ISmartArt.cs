namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a SmartArt graphic.
/// </summary>
public interface ISmartArt
{
    /// <summary>
    ///     Gets the collection of nodes in the SmartArt graphic.
    /// </summary>
    ISmartArtNodeCollection Nodes { get; }
}

internal sealed class SmartArt : ISmartArt
{
    internal SmartArt(SmartArtNodeCollection nodeCollection)
    {
        this.Nodes = nodeCollection;
    }

    public ISmartArtNodeCollection Nodes { get; }
}