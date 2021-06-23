using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    internal class LayoutChart : LayoutShape, IShape
    {
        public LayoutChart(SCSlideLayout slideLayout, P.GraphicFrame pGraphicFrame)
            : base(slideLayout, pGraphicFrame)
        {
        }

        public override SCSlideMaster ParentSlideMaster { get; }
    }
}