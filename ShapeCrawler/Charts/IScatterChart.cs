using System;
using ShapeCrawler.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    public interface IScatterChart : IChart
    {
    }
    
    internal class SCScatterChart : SCChart, IScatterChart
    {
        internal SCScatterChart(P.GraphicFrame pGraphicFrame, SCSlide slide)
            : base(pGraphicFrame, slide)
        {
            throw new NotImplementedException();
        }
    }
}