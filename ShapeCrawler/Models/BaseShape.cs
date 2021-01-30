using DocumentFormat.OpenXml;
using ShapeCrawler.Enums;
using ShapeCrawler.Extensions;

namespace ShapeCrawler.Models
{
    public interface IShape
    {
        uint Id { get; }

        long X { get; set; }

        long Y { get; set; }

        long Width { get; set; }

        long Height { get; }

        GeometryType GeometryType { get; }

        Placeholder Placeholder { get; }
    }

    public abstract class BaseShape
    {
        protected ISlide Slide { get; }

        internal readonly OpenXmlCompositeElement CompositeElement;

        public uint Id => CompositeElement.GetNonVisualDrawingProperties().Id;

        public abstract long X { get; }

        public abstract long Y { get; }

        public abstract long Width { get; }

        public abstract long Height { get; }
        public abstract GeometryType GeometryType { get; }

        #region Constructors

        protected BaseShape()
        {
            // TODO: Can be removed this parameterless constructor?
        }

        protected BaseShape(ISlide slide, OpenXmlCompositeElement compositeElement)
        {
            Slide = slide;
            CompositeElement = compositeElement;
        }

        #endregion
    }
}