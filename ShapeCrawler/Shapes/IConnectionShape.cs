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
        public SCConnectionShape(OpenXmlCompositeElement childOfPShapeTree, SCSlide slide)
            : base(childOfPShapeTree, slide)
        {
        }

        public SCShapeType ShapeType => SCShapeType.ConnectionShape;
    }
}