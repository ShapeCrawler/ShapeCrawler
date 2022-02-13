using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a connection shape.
    /// </summary>
    public interface IConnectionShape : IShape
    {
    }

    internal class SCConnectionShape : SlideShape, IConnectionShape
    {
        public SCConnectionShape(OpenXmlCompositeElement childOfpShapeTree, SCSlide slide)
            : base(childOfpShapeTree, slide)
        {
        }

        public ShapeType ShapeType => ShapeType.ConnectionShape;
    }
}