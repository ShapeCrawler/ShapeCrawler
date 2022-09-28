using ShapeCrawler.Charts;
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
    internal SCScatterChart(P.GraphicFrame pGraphicFrame, SCSlide slide)
        : base(pGraphicFrame, slide)
    {
    }
}