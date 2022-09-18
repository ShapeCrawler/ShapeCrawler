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
        ///     Gets shape fill.
        /// </summary>
        IShapeFill Fill { get; }

        /// <summary>
        ///     Gets text frame. Returns <c>NULL</c> if the AutoShape type doesn't have text frame.
        /// </summary>
        ITextFrame? TextFrame { get; }
    }
}