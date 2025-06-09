using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;

namespace ShapeCrawler.Charts;

internal sealed class AxisChart : Chart
{
    
    internal AxisChart( 
        XAxis xAxis, 
        SeriesCollection seriesCollection,
        SlideShapeOutline outline,
        ShapeFill fill,
        ChartPart chartPart) : base(seriesCollection, outline, fill, chartPart)
    {
        this.XAxis = xAxis;
    }
    
    public override IXAxis XAxis { get; }
}
