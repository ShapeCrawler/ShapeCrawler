﻿using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
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
}