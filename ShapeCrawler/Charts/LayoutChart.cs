using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    internal class LayoutChart : LayoutShape, IShape
    {
        public LayoutChart(SlideLayoutSc slideLayout, P.GraphicFrame pGraphicFrame) : base(slideLayout, pGraphicFrame)
        {
        }

        public override ThemePart ThemePart { get; }
        public override PresentationSc Presentation { get; }
        public override SlideMasterSc SlideMaster { get; }

        public string Name { get; }
        public bool Hidden { get; }
        public IPlaceholder Placeholder { get; }
        public string CustomData { get; set; }
    }
}