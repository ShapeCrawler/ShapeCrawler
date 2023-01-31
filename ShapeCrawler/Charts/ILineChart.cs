using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

/// <summary>
///     Represents a Line Chart.
/// </summary>
public interface ILineChart : IChart
{
}

internal sealed class SCLineChart : SCChart, ILineChart
{
    internal SCLineChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pGraphicFrame, parentSlideObject, parentShapeCollection)
    {
    }
}