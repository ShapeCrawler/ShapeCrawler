using System.Collections.Generic;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a collection of SmartArt nodes.
/// </summary>
public interface ISmartArtNodeCollection : IEnumerable<ISmartArtNode>
{
    /// <summary>
    ///     Gets the number of nodes in the collection.
    /// </summary>
    int Count { get; }
    
    /// <summary>
    ///     Adds a new node to the SmartArt graphic with the specified text.
    /// </summary>
    /// <param name="text">The text for the new node.</param>
    /// <returns>The newly added SmartArt node.</returns>
    ISmartArtNode AddNode(string text);
}
