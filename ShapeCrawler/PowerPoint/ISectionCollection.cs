using System.Collections.Generic;

namespace ShapeCrawler
{
    public interface ISectionCollection : IReadOnlyCollection<ISection>
    {
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
}