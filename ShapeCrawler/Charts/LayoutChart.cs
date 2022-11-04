using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

internal class LayoutChart : LayoutShape, IShape
{
    public LayoutChart(SCSlideLayout slideLayoutInternal, P.GraphicFrame pGraphicFrame)
        : base(slideLayoutInternal, pGraphicFrame)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Chart;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}