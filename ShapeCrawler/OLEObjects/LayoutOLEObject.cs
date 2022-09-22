using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable InconsistentNaming
namespace ShapeCrawler.OLEObjects
{
    /// <summary>
    ///     Represents a OLE Object on a Slide Layout.
    /// </summary>
    internal class LayoutOLEObject : LayoutShape, IShape
    {
        internal LayoutOLEObject(SCSlideLayout slideLayoutInternal, P.GraphicFrame pGraphicFrame)
            : base(slideLayoutInternal, pGraphicFrame)
        {
        }

        public SCShapeType ShapeType => SCShapeType.OLEObject;
    }
}