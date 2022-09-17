using System.Collections.Generic;
using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents paragraph collection.
    /// </summary>
    public interface IParagraphCollection : IReadOnlyList<IParagraph>
    {
        /// <summary>
        ///     Adds a new paragraph in collection.
        /// </summary>
        /// <returns>Added <see cref="SCParagraph" /> instance.</returns>
        IParagraph Add();
        
        /// <summary>
        ///     Removes specified paragraphs from collection.
        /// </summary>
        void Remove(IEnumerable<IParagraph> removeParagraphs);
    }
}