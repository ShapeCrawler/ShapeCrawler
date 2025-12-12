using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Presentations;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Charts;

internal sealed class ChartCollection(SlidePart slidePart)
{
    internal void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        this.AddPieChart(x, y, width, height, categoryValues, seriesName, "Pie Chart");
    }

    internal void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName,
        string chartName)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new PieChart(chartPart, categoryValues, seriesName).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, chartName);
    }

    internal void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new BarChart(chartPart, categoryValues, seriesName).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Bar Chart");
    }

    internal void AddScatterChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<double, double> pointValues,
        string seriesName)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new ScatterChart(chartPart, pointValues, seriesName).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Scatter Chart");
    }

    internal void AddStackedColumnChart(
        int x,
        int y,
        int width,
        int height,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new StackedColumnChart(chartPart, categoryValues, seriesNames).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Stacked Column Chart");
    }

    internal void AddClusteredBarChart(
        int x,
        int y,
        int width,
        int height,
        IList<string> categories,
        IList<DraftChart.SeriesData> seriesData,
        string chartName)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        var seriesTuples = seriesData.Select(s => (s.Name, s.Values)).ToList();
        new ClusteredBarChart(chartPart, categories, seriesTuples).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, chartName);
    }

    private void InsertChartGraphicFrame(ChartPart chartPart, int x, int y, int width, int height, string chartName)
    {
        var graphicFrame = new GraphicFrame
        {
            NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = this.GetNextShapeId(), Name = chartName },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            Transform = new Transform(
                new A.Offset { X = new Points(x).AsEmus(), Y = new Points(y).AsEmus() },
                new A.Extents { Cx = new Points(width).AsEmus(), Cy = new Points(height).AsEmus() }
            ),
            Graphic = new A.Graphic(
                new A.GraphicData(
                    new ChartReference { Id = slidePart.GetIdOfPart(chartPart) })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }
            )
        };

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    private uint GetNextShapeId()
    {
        var shapeIds = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0)
            .ToList();

        return shapeIds.Count > 0 ? shapeIds.Max() + 1 : 1;
    }
}