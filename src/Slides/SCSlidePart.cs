using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Presentations;
using ShapeCrawler.Units;

namespace ShapeCrawler.Slides;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

internal readonly ref struct SCSlidePart(SlidePart slidePart)
{
    internal void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        var rId = new SCOpenXmlPart(slidePart).GetNextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        GeneratePieChartContent(chartPart, categoryValues, seriesName);
        this.InsertChartGraphicFrame(chartPart, x, y, width, height);
    }

    internal void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        var rId = new SCOpenXmlPart(slidePart).GetNextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        GenerateBarChartContent(chartPart, categoryValues, seriesName);
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Bar Chart");
    }

    private static void GeneratePieChartContent(
        ChartPart chartPart,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        var chartSpace = new ChartSpace();
        chartSpace.Append(new EditingLanguage { Val = "en-US" });
        chartSpace.Append(new RoundedCorners { Val = false });

        var chart = new Chart();
        chart.Append(new AutoTitleDeleted { Val = false });

        var plotArea = new PlotArea();
        plotArea.Append(new Layout());

        var pieChart = new PieChart();
        pieChart.Append(new VaryColors { Val = true });

        var series = new PieChartSeries(
            new Index { Val = 0 },
            new Order { Val = 0 },
            new SeriesText(new NumericValue { Text = seriesName }));

        // --- Categories ---
        var categoriesCount = UInt32Value.FromUInt32((uint)categoryValues.Count);
        var stringLiteral = new StringLiteral { PointCount = new PointCount { Val = categoriesCount } };
        uint catIndex = 0;
        foreach (var categoryToValue in categoryValues)
        {
            stringLiteral.Append(new StringPoint
            {
                Index = catIndex,
                NumericValue = new NumericValue(categoryToValue.Key)
            });
            catIndex++;
        }

        var categoryAxisData = new CategoryAxisData();
        categoryAxisData.Append(stringLiteral);

        // --- Values ---
        var numberLiteral = new NumberLiteral
        {
            FormatCode = new FormatCode("General"),
            PointCount = new PointCount { Val = categoriesCount }
        };
        catIndex = 0;
        foreach (var categoryToValue in categoryValues)
        {
            numberLiteral.Append(new NumericPoint
            {
                Index = catIndex,
                NumericValue = new NumericValue(categoryToValue.Value.ToString())
            });
            catIndex++;
        }

        var values = new Values();
        values.Append(numberLiteral);

        // Append categories and values to the series
        series.Append(categoryAxisData);
        series.Append(values);

        pieChart.Append(series);
        plotArea.Append(pieChart);

        chart.Append(plotArea);

        var cLegend = new Legend(
            new LegendPosition { Val = LegendPositionValues.Right });
        chart.Append(cLegend);

        chartSpace.Append(chart);
        chartPart.ChartSpace = chartSpace;

        // Show Data Labels
        pieChart.Append(
            new DataLabels(
                new ShowLegendKey { Val = false },
                new ShowValue { Val = true },
                new ShowCategoryName { Val = false },
                new ShowSeriesName { Val = false },
                new ShowPercent { Val = false },
                new ShowBubbleSize { Val = false },
                new ShowLeaderLines { Val = true }));
    }

    private static void GenerateBarChartContent(
        ChartPart chartPart,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        // Create the ChartSpace element
        var chartSpace = new ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        
        // Basic chart settings
        chartSpace.Append(new EditingLanguage { Val = "en-US" });
        chartSpace.Append(new RoundedCorners { Val = false });

        var chart = new Chart();
        chart.Append(new AutoTitleDeleted { Val = false });
        
        // Create plot area
        var plotArea = new PlotArea();
        plotArea.Append(new Layout());

        // Create bar chart
        var barChart = new BarChart();
        barChart.Append(new BarDirection { Val = BarDirectionValues.Column });
        barChart.Append(new BarGrouping { Val = BarGroupingValues.Clustered });
        barChart.Append(new VaryColors { Val = false });

        // Create series
        var series = new BarChartSeries();
        series.Append(new Index { Val = 0 });
        series.Append(new Order { Val = 0 });
        
        // Add series name
        var seriesText = new SeriesText(new NumericValue { Text = seriesName });
        series.Append(seriesText);

        // --- Categories ---
        var categoriesCount = UInt32Value.FromUInt32((uint)categoryValues.Count);
        var categoryAxisData = new CategoryAxisData();
        var stringLiteral = new StringLiteral();
        stringLiteral.Append(new PointCount { Val = categoriesCount });
        
        uint index = 0;
        foreach (var item in categoryValues)
        {
            var point = new StringPoint { Index = index };
            point.Append(new NumericValue(item.Key));
            stringLiteral.Append(point);
            index++;
        }
        
        categoryAxisData.Append(stringLiteral);
        series.Append(categoryAxisData);

        // --- Values ---
        var values = new Values();
        var numberLiteral = new NumberLiteral();
        numberLiteral.Append(new FormatCode("General"));
        numberLiteral.Append(new PointCount { Val = categoriesCount });
        
        index = 0;
        foreach (var item in categoryValues)
        {
            var point = new NumericPoint { Index = index };
            point.Append(new NumericValue(item.Value.ToString()));
            numberLiteral.Append(point);
            index++;
        }
        
        values.Append(numberLiteral);
        series.Append(values);

        // Add the series to the bar chart
        barChart.Append(series);

        // Add axis IDs
        uint axisId1 = 1U;
        uint axisId2 = 2U;
        barChart.Append(new AxisId { Val = axisId1 });
        barChart.Append(new AxisId { Val = axisId2 });

        // Add the bar chart to the plot area
        plotArea.Append(barChart);

        // Add category axis
        var categoryAxis = new CategoryAxis();
        categoryAxis.Append(new AxisId { Val = axisId1 });
        
        var scaling = new Scaling();
        scaling.Append(new Orientation { Val = OrientationValues.MinMax });
        categoryAxis.Append(scaling);
        
        categoryAxis.Append(new Delete { Val = false });
        categoryAxis.Append(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.Append(new CrossingAxis { Val = axisId2 });
        
        // Add the category axis to the plot area
        plotArea.Append(categoryAxis);

        // Add value axis
        var valueAxis = new ValueAxis();
        valueAxis.Append(new AxisId { Val = axisId2 });
        
        scaling = new Scaling();
        scaling.Append(new Orientation { Val = OrientationValues.MinMax });
        valueAxis.Append(scaling);
        
        valueAxis.Append(new Delete { Val = false });
        valueAxis.Append(new AxisPosition { Val = AxisPositionValues.Left });
        valueAxis.Append(new CrossingAxis { Val = axisId1 });
        
        // Add the value axis to the plot area
        plotArea.Append(valueAxis);
        
        // Add the plot area to the chart
        chart.Append(plotArea);

        // Add legend
        var legend = new Legend();
        legend.Append(new LegendPosition { Val = LegendPositionValues.Right });
        
        // Add the legend to the chart
        chart.Append(legend);
        
        // Add the chart to the chart space
        chartSpace.Append(chart);
        
        // Save the chart part
        chartPart.ChartSpace = chartSpace;
    }

    private void InsertChartGraphicFrame(ChartPart chartPart, int x, int y, int width, int height)
    {
        // Create a new GraphicFrame
        var graphicFrame = new GraphicFrame
        {
            // Give it an ID and name
            NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = this.GetNextShapeId(), Name = "Pie Chart" },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),

            Transform = new Transform(
                new A.Offset { X = new Pixels(x).AsHorizontalEmus(), Y = new Pixels(y).AsVerticalEmus() },
                new A.Extents { Cx = new Pixels(width).AsHorizontalEmus(), Cy = new Pixels(height).AsVerticalEmus() }),
            Graphic = new A.Graphic(
                new A.GraphicData(
                        new ChartReference { Id = slidePart.GetIdOfPart(chartPart) })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })
        };

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    private void InsertChartGraphicFrame(ChartPart chartPart, int x, int y, int width, int height, string chartName)
    {
        // Create a new GraphicFrame
        var graphicFrame = new GraphicFrame
        {
            // Give it an ID and name
            NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = this.GetNextShapeId(), Name = chartName },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),

            Transform = new Transform(
                new A.Offset { X = new Pixels(x).AsHorizontalEmus(), Y = new Pixels(y).AsVerticalEmus() },
                new A.Extents { Cx = new Pixels(width).AsHorizontalEmus(), Cy = new Pixels(height).AsVerticalEmus() }),
            Graphic = new A.Graphic(
                new A.GraphicData(
                        new ChartReference { Id = slidePart.GetIdOfPart(chartPart) })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })
        };

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    private uint GetNextShapeId()
    {
        // Get all existing shape IDs from the slide
        var shapeIds = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0)
            .ToList();
        
        // Find the maximum ID and add 1, or start with 1 if no shapes exist
        return (uint)(shapeIds.Count > 0 ? shapeIds.Max() + 1 : 1);
    }
}