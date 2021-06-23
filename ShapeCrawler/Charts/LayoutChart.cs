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
            // TODO: add test for reading chart placeholder on Layout
        }

        public override SCSlideMaster ParentSlideMaster { get; }
    }
}