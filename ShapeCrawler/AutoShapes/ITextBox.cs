using ShapeCrawler.Collections;

namespace ShapeCrawler.AutoShapes
{
    /// <summary>
    ///     Represents a text box.
    /// </summary>
    public interface ITextBox // TODO: rename to TextFrame?
    {
        /// <summary>
        ///     Gets text paragraph collection.
        /// </summary>
        IParagraphCollection Paragraphs { get; }

        /// <summary>
        ///     Gets or sets text box string content. Returns <c>NULL</c> if the text box is empty.
        ///     <para>NOTE: setter removes all paragraphs except first, which will be used as the single paragraph to set box content.</para>
        /// </summary>
        string Text { get; set; }

        /// <summary>
        ///     Gets text fit type.
        /// </summary>
        SCAutoFitType SCAutoFitType { get; }
    }
}