using System.Collections.Generic;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collections of Slide Masters.
/// </summary>
public interface ISlideMasterCollection
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets the element at the specified index.
    /// </summary>
    ISlideMaster this[int index] { get; }

    /// <summary>
    ///     Gets the generic enumerator that iterates through the collection.
    /// </summary>
    IEnumerator<ISlideMaster> GetEnumerator();
}