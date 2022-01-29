using ShapeCrawler.Collections;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    public interface ILineChart : IChart
    {
        public ICategoryCollection Categories { get; }
    }

    internal class SCLineChart : SCChart, ILineChart
    {
        internal SCLineChart(P.GraphicFrame pGraphicFrame, SCSlide slide)
            : base(pGraphicFrame, slide)
        {
        }
    }
}