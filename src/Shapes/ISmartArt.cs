namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a SmartArt graphic.
/// </summary>
public interface ISmartArt : IShape
{
    /// <summary>
    ///     Gets the collection of nodes in the SmartArt graphic.
    /// </summary>
    ISmartArtNodeCollection Nodes { get; }
}
