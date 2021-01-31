using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;

namespace ShapeCrawler.Models
{
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