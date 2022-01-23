using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Factories
{
    internal class ChartGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        internal override IShape? Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide, SlideGroupShape groupShape)
        {
            if (pShapeTreeChild is not P.GraphicFrame pGraphicFrame)
            {
                return this.Successor?.Create(pShapeTreeChild, slide, groupShape);
            }

            var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>();
            if (aGraphicData!.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                // Get chart part
                var cChartReference = pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>()
                    .GetFirstChild<C.ChartReference>();

                // var slide = this.ParentSlideLayoutInternal;
                var sdkChartPart = (ChartPart)slide.SlidePart.GetPartById(cChartReference.Id);

                C.PlotArea cPlotArea = sdkChartPart.ChartSpace.GetFirstChild<C.Chart>().PlotArea;
                var cCharts = cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
                
                
                var chart = new SCChart(pGraphicFrame, slide);

                return chart;
            }

            return Successor?.Create(pShapeTreeChild, slide, groupShape);
        }
    }
}