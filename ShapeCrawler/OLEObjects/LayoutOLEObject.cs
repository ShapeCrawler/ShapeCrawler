using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.OLEObjects
{
    internal class LayoutOLEObject : LayoutShape, IShape
    {
        public LayoutOLEObject(SlideLayoutSc slideLayout, P.GraphicFrame pGraphicFrame) : base(slideLayout,
            pGraphicFrame)
        {
        }

        public string Name { get; }
        public bool Hidden { get; }
        public IPlaceholder Placeholder { get; }
        public string CustomData { get; set; }
    }
}