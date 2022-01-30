using DocumentFormat.OpenXml;

namespace ShapeCrawler.Shapes
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
    }
}