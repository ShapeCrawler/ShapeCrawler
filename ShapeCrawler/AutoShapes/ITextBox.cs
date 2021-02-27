using ShapeCrawler.Texts;

namespace ShapeCrawler.AutoShapes
{
    public interface ITextBox
    {
        /// <summary>
        ///     Gets text paragraph collection.
        /// </summary>
        ParagraphCollection Paragraphs { get; }

        /// <summary>
        ///     Gets or sets text box string content. Returns <c>NULL</c> if the text box is empty.
        /// </summary>
        string Text { get; set; }
    }
}