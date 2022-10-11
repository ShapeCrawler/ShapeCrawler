using ShapeCrawler.SlideMasters;
using OneOf;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents Pie chart interface.
/// </summary>
public interface IPieChart : IChart
{
}

internal sealed class SCPieChart : SCChart, IPieChart
{
    internal SCPieChart(DocumentFormat.OpenXml.Presentation.GraphicFrame pGraphicFrame, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
        : base(pGraphicFrame, oneOfSlide)
    {
    }
}