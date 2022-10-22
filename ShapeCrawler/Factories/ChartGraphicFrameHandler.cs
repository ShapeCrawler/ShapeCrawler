using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using OneOf;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories;

internal class ChartGraphicFrameHandler : OpenXmlElementHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    
    internal override Shape? Create(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, SCGroupShape groupShape)
    {
        if (pShapeTreeChild is not P.GraphicFrame pGraphicFrame)
        {
            return this.Successor?.Create(pShapeTreeChild, slideObject, groupShape);
        }

        var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
        if (!aGraphicData.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
        {
            return this.Successor?.Create(pShapeTreeChild, slideObject, groupShape);
        }

        var slideBase = slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
        var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>()!;
        var chartPart = (ChartPart)slideBase.TypedOpenXmlPart.GetPartById(cChartRef.Id!);
        var cPlotArea = chartPart!.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea;
        var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        if (cCharts.Count() > 1)
        {
            return new SCComboChart(pGraphicFrame, slideObject);
        }
                
        var chartTypeName = cCharts.Single().LocalName;
                
        if (chartTypeName == "lineChart")
        {
            return new SCLineChart(pGraphicFrame, slideObject);
        }
                
        if (chartTypeName == "barChart")
        {
            return new SCBarChart(pGraphicFrame, slideObject);
        }
    
        if (chartTypeName == "pieChart")
        {
            return new SCPieChart(pGraphicFrame, slideObject);
        }
                
        if (chartTypeName == "scatterChart")
        {
            return new SCScatterChart(pGraphicFrame, slideObject);
        }

        return new SCChart(pGraphicFrame, slideObject);
    }
}