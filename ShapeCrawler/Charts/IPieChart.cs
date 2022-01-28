using System;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    public interface IPieChart : IChart
    {
        public ICategoryCollection Categories { get; }
    }
    
    internal sealed class SCPieChart : SCChart, IPieChart
    {
        internal SCPieChart(DocumentFormat.OpenXml.Presentation.GraphicFrame pGraphicFrame, SCSlide slide)
            : base(pGraphicFrame, slide)
        {
            throw new NotImplementedException();
        }
    }
}