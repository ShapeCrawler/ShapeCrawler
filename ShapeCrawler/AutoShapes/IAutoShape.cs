using ShapeCrawler.Models;
using ShapeCrawler.Models.Styles;
using ShapeCrawler.Texts;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an auto shape on a slide.
    /// </summary>
    public interface IAutoShape : IShape
    {
        /// <summary>
        ///     Gets shape fill object that contains fill formatting properties for the shape.
        ///     Returns <c>NULL</c> if the shape is not filled.
        /// </summary>
        ShapeFill Fill { get; }

        /// <summary>
        ///     Gets text box. Returns <c>NULL</c> if shape has not text content.
        /// </summary>
        TextBoxSc TextBox { get; }
    }
}