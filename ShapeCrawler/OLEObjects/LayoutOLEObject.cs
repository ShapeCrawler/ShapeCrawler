using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.OLEObjects
{
    /// <summary>
    ///     Represents a OLE Object on a Slide Layout.
    /// </summary>
    internal class LayoutOLEObject : LayoutShape, IShape
    {
        internal LayoutOLEObject(SCSlideLayout slideLayout, P.GraphicFrame pGraphicFrame)
            : base(slideLayout, pGraphicFrame)
        {
        }

        public override SCSlideMaster ParentSlideMaster { get; }
    }
}