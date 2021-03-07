using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;

namespace ShapeCrawler.Experiment
{
    public abstract class BaseShape
    {
        protected ISlide Slide { get; }

        internal OpenXmlCompositeElement CompositeElement { get; }

        public uint Id => CompositeElement.NonVisualDrawingProperties().Id;

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

        protected BaseShape(ISlide slide, OpenXmlCompositeElement shapeTreeSource)
        {
            Slide = slide;
            CompositeElement = shapeTreeSource;
        }

        #endregion
    }
}