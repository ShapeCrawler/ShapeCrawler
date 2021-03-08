using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.OLEObjects
{
    /// <summary>
    ///     Represents a OLE Object on a Slide Layout.
    /// </summary>
    internal class LayoutOLEObject : LayoutShape, IShape
    {
        internal LayoutOLEObject(SlideLayoutSc slideLayout, P.GraphicFrame pGraphicFrame) 
            : base(slideLayout, pGraphicFrame)
        {
        }
    }
}