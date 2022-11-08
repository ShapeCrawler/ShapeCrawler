using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

/// <summary>
///     Represents a Scatter Chart.
/// </summary>
public interface IScatterChart : IChart
{
}

internal class SCScatterChart : SCChart, IScatterChart
{
    internal SCScatterChart(P.GraphicFrame pGraphicFrame, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
        : base(pGraphicFrame, oneOfSlide)
    {
    }
}