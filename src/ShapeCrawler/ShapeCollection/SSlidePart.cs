using System.Collections.Generic;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.ShapeCollection;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

internal readonly ref struct SSlidePart(SlidePart slidePart)
{
    internal void AddPieChart(int x, int y, int width, int height, Dictionary<string, double> categoryValues, string seriesName)
    {
        var rId = new SOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        GeneratePieChartContent(chartPart, categoryValues, seriesName);
        this.InsertChartGraphicFrame(chartPart);
    }
    
    private static void GeneratePieChartContent(ChartPart chartPart, Dictionary<string, double> categoryValues, string seriesName)
    {
        var chartSpace = new ChartSpace();
        chartSpace.Append(new EditingLanguage { Val = "en-US" });
        chartSpace.Append(new RoundedCorners { Val = false });
        
        var chart = new Chart();
        chart.Append(new AutoTitleDeleted { Val = false });

        PlotArea plotArea = new PlotArea();
        plotArea.Append(new Layout());
        
        var pieChart = new PieChart();
        pieChart.Append(new VaryColors { Val = true });
        
        PieChartSeries series = new PieChartSeries(
            new Index { Val = 0 },
            new Order { Val = 0 },
            new SeriesText(new NumericValue { Text = seriesName }));

        // --- Categories ---
        var categoriesCount = UInt32Value.FromUInt32((uint)categoryValues.Count);
        var stringLiteral = new StringLiteral
        {
            PointCount = new PointCount { Val = categoriesCount }
        };
        uint catIndex = 0;
        foreach (var categoryToValue in categoryValues)
        {
            stringLiteral.Append(new StringPoint { Index = catIndex, NumericValue = new NumericValue(categoryToValue.Key) });    
            catIndex++;
        }

        CategoryAxisData categoryAxisData = new CategoryAxisData();
        categoryAxisData.Append(stringLiteral);

        // --- Values ---
        var numberLiteral = new NumberLiteral { FormatCode = new FormatCode("General"), PointCount = new PointCount { Val = categoriesCount } };
        catIndex = 0;
        foreach (var categoryToValue in categoryValues)
        {
            numberLiteral.Append(new NumericPoint { Index = catIndex, NumericValue = new NumericValue(categoryToValue.Value.ToString()) });
            catIndex++;
        }

        Values values = new Values();
        values.Append(numberLiteral);

        // Append categories and values to the series
        series.Append(categoryAxisData);
        series.Append(values);

        pieChart.Append(series);
        plotArea.Append(pieChart);
        
        chart.Append(plotArea);
        
        var cLegendPos = new C.LegendPosition { Val = C.LegendPositionValues.Right };
        var cLegend = new C.Legend();
        cLegend.Append(cLegendPos);
        chart.Append(cLegend);
        
        chartSpace.Append(chart);
        chartPart.ChartSpace = chartSpace;

        // Show Data Labels
        var cShowLegendKey = new C.ShowLegendKey { Val = false };
        var cShowVal = new C.ShowValue { Val = true };
        var cShowCatName = new ShowCategoryName{ Val = false };
        var cShowSerName = new ShowSeriesName{ Val = false };
        var cShowPercent = new ShowPercent{ Val = false };
        var cShowBubbleSize = new ShowBubbleSize{ Val = false };
        var cShowLeaderLines = new ShowLeaderLines{ Val = true };
        var cdLbls = new C.DataLabels();
        cdLbls.Append(cShowLegendKey);
        cdLbls.Append(cShowVal);
        cdLbls.Append(cShowCatName);
        cdLbls.Append(cShowSerName);
        cdLbls.Append(cShowPercent);
        cdLbls.Append(cShowBubbleSize);
        cdLbls.Append(cShowLeaderLines);
        // series.Append(cdLbls);
        pieChart.Append(cdLbls);
    }

    /// <summary>
    /// Inserts a graphic frame in the slide that references the chart part by ID.
    /// </summary>
    private void InsertChartGraphicFrame(ChartPart chartPart)
    {
        // Retrieve the shape tree of the slide
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree;

        // Create a new GraphicFrame
        var graphicFrame = new GraphicFrame
        {
            // Give it an ID and name
            NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties() { Id = 2U, Name = "PieChart" },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            
            // Position/Size of the chart in EMUs (English Metric Units)
            // For reference: 914400 EMUs = 1 inch
            Transform = new Transform(
                new A.Offset() { X = 1524000L, Y = 1524000L }, // 1.67in from top-left corner
                new A.Extents() { Cx = 6096000L, Cy = 3429000L }), // 6.67in wide, 3.75in high
            Graphic = new A.Graphic(
                new A.GraphicData(
                    new ChartReference() { Id = slidePart.GetIdOfPart(chartPart) }) // "rId3"
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
        };

        // Append the graphic frame to the shape tree
        shapeTree!.Append(graphicFrame);
    }
}