using ShapeCrawler.SmartArts;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

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