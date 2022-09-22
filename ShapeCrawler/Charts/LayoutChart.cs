using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    internal class LayoutChart : LayoutShape, IShape
    {
        public LayoutChart(SCSlideLayout slideLayoutInternal, P.GraphicFrame pGraphicFrame)
            : base(slideLayoutInternal, pGraphicFrame)
        {
            // TODO: add test for reading chart placeholder on Layout
        }

        public SCShapeType ShapeType => SCShapeType.Chart;
    }
}