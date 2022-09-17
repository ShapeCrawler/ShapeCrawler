using ShapeCrawler.Collections;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a Line Chart.
    /// </summary>
    public interface ILineChart : IChart
    {
    }

    internal class SCLineChart : SCChart, ILineChart
    {
        internal SCLineChart(P.GraphicFrame pGraphicFrame, SCSlide slide)
            : base(pGraphicFrame, slide)
        {
        }
    }
}