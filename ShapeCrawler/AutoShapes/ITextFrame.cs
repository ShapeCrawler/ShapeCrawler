using ShapeCrawler.Collections;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents text frame.
    /// </summary>
    public interface ITextFrame
    {
        /// <summary>
        ///     Gets collection of paragraphs.
        /// </summary>
        IParagraphCollection Paragraphs { get; }

        /// <summary>
        ///     Gets or sets text.
        /// </summary>
        string Text { get; set; }

        /// <summary>
        ///     Gets AutoFit type.
        /// </summary>
        SCAutoFitType AutoFitType { get; }
        
        /// <summary>
        ///     Gets a value indicating whether text frame can be changed.
        /// </summary>
        bool CanChange { get; }
    }
}