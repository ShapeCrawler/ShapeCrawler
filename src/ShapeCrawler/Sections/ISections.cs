using System.Collections.Generic;

namespace ShapeCrawler;

/// <summary>
///     Represents collection of presentation section.
/// </summary>
public interface ISections : IReadOnlyCollection<ISection>
{
    /// <summary>
    ///     Gets section by index.
    /// </summary>
    ISection this[int index] { get; }

    /// <summary>
    ///     Removes specified section.
    /// </summary>
    void Remove(ISection removingSection);

    /// <summary>
    ///     Gets section by section name.
    /// </summary>
    ISection GetByName(string sectionName);
}