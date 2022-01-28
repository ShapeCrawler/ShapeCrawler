using System;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents a Bar or Column chart.
    /// </summary>
    public interface IBarChart : IChart
    {
        public ICategoryCollection Categories { get; }
    }
    
    internal sealed class SCBarChart : SCChart, IBarChart
    {
        internal SCBarChart(P.GraphicFrame pGraphicFrame, SCSlide slide)
            : base(pGraphicFrame, slide)
        {
            throw new NotImplementedException();
        }
    }
}