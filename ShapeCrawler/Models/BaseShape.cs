using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Enums;

namespace ShapeCrawler.Models
{
    public abstract class BaseShape
    {
#if NETCOREAPP2_0 || NETSTANDARD2_0 || NETSTANDARD2_1
        internal readonly OpenXmlCompositeElement _compositeElement;
#else
        protected internal readonly OpenXmlCompositeElement _compositeElement;
#endif

        public uint Id => _compositeElement.GetFirstChild<NonVisualShapeProperties>().NonVisualDrawingProperties.Id;

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

        protected BaseShape(OpenXmlCompositeElement compositeElement)
        {
            _compositeElement = compositeElement;
        }

        #endregion
    }
}