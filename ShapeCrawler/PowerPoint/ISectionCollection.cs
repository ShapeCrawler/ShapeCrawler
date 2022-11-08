using System.Collections.Generic;

namespace ShapeCrawler;

/// <summary>
///     Represents collection of presentation section.
/// </summary>
public interface ISectionCollection : IReadOnlyCollection<ISection>
{
    /// <summary>
    ///     Gets section by index.
    /// </summary>
    ISection this[int i] { get; }

    /// <summary>
    ///     Removes specified section.
    /// </summary>
    void Remove(ISection removingSection);

    /// <summary>
    ///     Gets section by section name.
    /// </summary>
    ISection GetByName(string sectionName);
}