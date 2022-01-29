using System;
using ShapeCrawler.Collections;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
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
        }
    }
}