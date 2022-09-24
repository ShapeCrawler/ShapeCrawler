﻿using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart on a Slide Master.
    /// </summary>
    internal class MasterChart : MasterShape, IShape
    {
        internal MasterChart(SCSlideMaster slideMasterInternal, P.GraphicFrame pGraphicFrame)
            : base(pGraphicFrame, slideMasterInternal)
        {
        }

        public ShapeType ShapeType => ShapeType.Chart;
        public override SCPresentation PresentationInternal { get; }
    }
}