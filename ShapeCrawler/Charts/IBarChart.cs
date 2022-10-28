using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;
using OneOf;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a Bar or Column chart.
    /// </summary>
    public interface IBarChart : IChart
    {
    }

    internal sealed class SCBarChart : SCChart, IBarChart
    {
        internal SCBarChart(P.GraphicFrame pGraphicFrame, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
            : base(pGraphicFrame, oneOfSlide)
        {
        }
    }
}