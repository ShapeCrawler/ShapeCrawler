using System.Collections.Generic;

namespace SlideXML.Models.TextBody
{
    /// <summary>
    /// Represents a text frame of the shape.
    /// </summary>
    public interface ITextFrame
    {
        /// <summary>
        /// Returns list of paragraphs.
        /// </summary>
        IList<Paragraph> Paragraphs { get; }

        /// <summary>
        /// Returns text content.
        /// </summary>
        string Text { get; }
    }
}