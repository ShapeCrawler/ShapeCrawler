using System.Collections.Generic;

namespace ShapeCrawler.Texts
{
    /// <summary>
    /// Represents a text frame of the shape.
    /// </summary>
    public interface ITextFrame
    {
        /// <summary>
        /// Returns list of paragraphs.
        /// </summary>
        IList<ParagraphSc> Paragraphs { get; }

        /// <summary>
        /// Returns text content.
        /// </summary>
        string Text { get; }
    }
}