using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents interface of AutoShape.
    /// </summary>
    public interface IAutoShape : IShape
    {
        /// <summary>
        ///     Gets shape fill object that contains fill formatting properties for the shape.
        /// </summary>
        IShapeFill Fill { get; }

        /// <summary>
        ///     Gets text box.
        /// </summary>
        ITextBox? TextBox { get; } // TODO: always return TextBox
    }
}