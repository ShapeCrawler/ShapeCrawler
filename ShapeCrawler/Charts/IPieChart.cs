using ShapeCrawler.Collections;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents Pie chart interface.
    /// </summary>
    public interface IPieChart : IChart
    {
    }

    internal sealed class SCPieChart : SCChart, IPieChart
    {
        internal SCPieChart(DocumentFormat.OpenXml.Presentation.GraphicFrame pGraphicFrame, SCSlide slide)
            : base(pGraphicFrame, slide)
        {
        }
    }
}