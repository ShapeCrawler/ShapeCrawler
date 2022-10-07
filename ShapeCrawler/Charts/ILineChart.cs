using ShapeCrawler.Charts;
using ShapeCrawler.SlideMasters;
using OneOf;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

/// <summary>
///     Represents a Line Chart.
/// </summary>
public interface ILineChart : IChart
{
}

internal class SCLineChart : SCChart, ILineChart
{
    internal SCLineChart(P.GraphicFrame pGraphicFrame, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject)
        : base(pGraphicFrame, slideObject)
    {
    }
}