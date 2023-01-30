using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents a Bar or Column chart.
/// </summary>
public interface IBarChart : IChart
{
}

internal sealed class SCBarChart : SCChart, IBarChart
{
    internal SCBarChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pGraphicFrame, parentSlideObject, parentShapeCollection)
    {
    }
}