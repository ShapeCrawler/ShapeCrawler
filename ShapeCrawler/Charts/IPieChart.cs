using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents Pie chart interface.
/// </summary>
public interface IPieChart : IChart
{
}

internal sealed class SCPieChart : SCChart, IPieChart
{
    internal SCPieChart(
        P.GraphicFrame pGraphicFrame, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pGraphicFrame, parentSlideObject, parentShapeCollection)
    {
    }
}