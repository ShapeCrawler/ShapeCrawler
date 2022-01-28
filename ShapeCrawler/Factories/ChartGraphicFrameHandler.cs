using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

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

            var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
            if (!aGraphicData.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                return this.Successor?.Create(pShapeTreeChild, slide, groupShape);
            }

            var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>();
            var chartPart = (ChartPart)slide.SlidePart.GetPartById(cChartRef.Id);
            var cPlotArea = chartPart!.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea;
            var cCharts = cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

            if (cCharts.Count() > 1)
            {
                return new SCComboChart(pGraphicFrame, slide);
            }
                
            var chartTypeName = cCharts.Single().LocalName;
                
            if (chartTypeName == "LineChart")
            {
                return new SCLineChart(pGraphicFrame, slide);
            }
                
            if (chartTypeName == "BarChart")
            {
                return new SCBarChart(pGraphicFrame, slide);
            }
    
            if (chartTypeName == "PieChart")
            {
                return new SCPieChart(pGraphicFrame, slide);
            }
                
            if (chartTypeName == "ScatterChart")
            {
                return new SCScatterChart(pGraphicFrame, slide);
            }

            return null; // rest of chart types are not supported yet

        }
    }

}