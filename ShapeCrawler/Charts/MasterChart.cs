using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents a chart on a Slide Master.
/// </summary>
internal class MasterChart : MasterShape, IShape
{
    internal MasterChart(SCSlideMaster slideMasterInternal, P.GraphicFrame pGraphicFrame)
        : base(pGraphicFrame, slideMasterInternal)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Chart;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}