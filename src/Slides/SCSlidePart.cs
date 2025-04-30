using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Slides;

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

    internal ISmartArt AddSmartArt(int x, int y, int width, int height, SmartArtType smartArtType)
    {
        // Create a new GraphicFrame
        var graphicFrame = new GraphicFrame();

        // Add ID and name properties
        var nvGraphicFrameProperties = new NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new NonVisualDrawingProperties
        {
            Id = this.GetNextShapeId(), Name = $"SmartArt {smartArtType}"
        };
        var nonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        nvGraphicFrameProperties.Append(nonVisualDrawingProperties);
        nvGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
        nvGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);
        graphicFrame.Append(nvGraphicFrameProperties);

        // Add transform properties
        var transform = new Transform();
        transform.Append(new A.Offset { X = new Pixels(x).AsHorizontalEmus(), Y = new Pixels(y).AsVerticalEmus() });
        transform.Append(new A.Extents
        {
            Cx = new Pixels(width).AsHorizontalEmus(), Cy = new Pixels(height).AsVerticalEmus()
        });
        graphicFrame.Append(transform);

        // Create the diagram graphic
        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" };

        // Instead of using Diagram class directly, we'll use a simple approach
        // with just a GraphicData container that identifies as a diagram
        // This will create a valid empty SmartArt shell that can be modified later
        graphic.Append(graphicData);
        graphicFrame.Append(graphic);

        // Add to slide
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);

        // Return a simplified SmartArt implementation for now
        return new SmartArt(graphicFrame);
    }

    private static void GeneratePieChartContent(
        ChartPart chartPart,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        var chart = new Chart();
        chart.AppendChild(new AutoTitleDeleted { Val = false });

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
                Index = catIndex, NumericValue = new NumericValue(categoryToValue.Key)
            });
            catIndex++;
        }

        var categoryAxisData = new CategoryAxisData();
        categoryAxisData.Append(stringLiteral);

        // --- Values ---
        var numberLiteral = new NumberLiteral
        {
            FormatCode = new FormatCode("General"), PointCount = new PointCount { Val = categoriesCount }
        };
        catIndex = 0;
        foreach (var categoryToValue in categoryValues)
        {
            numberLiteral.Append(new NumericPoint
            {
                Index = catIndex, NumericValue = new NumericValue(categoryToValue.Value.ToString())
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
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var chart = new Chart();
        chart.AppendChild(new AutoTitleDeleted { Val = false });

        // Create series
        var series = new BarChartSeries(new Index { Val = 0 }, new Order { Val = 0 });

        // Add series name
        var seriesText = new SeriesText();
        seriesText.AppendChild(new NumericValue { Text = seriesName });
        series.AppendChild(seriesText);

        // --- Categories ---
        var categoriesCount = UInt32Value.FromUInt32((uint)categoryValues.Count);
        var categoryAxisData = new CategoryAxisData();
        var stringLiteral = new StringLiteral(new PointCount { Val = categoriesCount });

        uint index = 0;
        foreach (var item in categoryValues)
        {
            var point = new StringPoint() { Index = index };
            point.AppendChild(new NumericValue(item.Key));
            stringLiteral.AppendChild(point);
            index++;
        }

        categoryAxisData.AppendChild(stringLiteral);
        series.AppendChild(categoryAxisData);

        // --- Values ---
        var values = new Values();
        var numberLiteral = new NumberLiteral(new FormatCode("General"), new PointCount { Val = categoriesCount });

        index = 0;
        foreach (var item in categoryValues)
        {
            var point = new NumericPoint { Index = index };
            point.AppendChild(new NumericValue(item.Value.ToString()));
            numberLiteral.AppendChild(point);
            index++;
        }

        values.AppendChild(numberLiteral);
        series.AppendChild(values);

        // Create bar chart
        const uint axisId1 = 1U;
        const uint axisId2 = 2U;

        // Add the bar chart to the plot area
        var plotArea = new PlotArea(
            new Layout(), new BarChart(
                new BarDirection { Val = BarDirectionValues.Column },
                new BarGrouping { Val = BarGroupingValues.Clustered },
                new VaryColors { Val = false },
                series,
                new AxisId { Val = axisId1 },
                new AxisId { Val = axisId2 }));

        // Add category axis
        var categoryAxis = new CategoryAxis();
        categoryAxis.AppendChild(new AxisId { Val = axisId1 });

        var scaling = new Scaling();
        scaling.AppendChild(new Orientation { Val = OrientationValues.MinMax });
        categoryAxis.AppendChild(scaling);

        categoryAxis.AppendChild(new Delete { Val = false });
        categoryAxis.AppendChild(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.AppendChild(new CrossingAxis { Val = axisId2 });

        // Add the category axis to the plot area
        plotArea.AppendChild(categoryAxis);

        // Add value axis
        var valueAxis = new ValueAxis();
        valueAxis.AppendChild(new AxisId { Val = axisId2 });

        scaling = new Scaling();
        scaling.AppendChild(new Orientation { Val = OrientationValues.MinMax });
        valueAxis.AppendChild(scaling);

        valueAxis.AppendChild(new Delete { Val = false });
        valueAxis.AppendChild(new AxisPosition { Val = AxisPositionValues.Left });
        valueAxis.AppendChild(new CrossingAxis { Val = axisId1 });

        // Add the value axis to the plot area
        plotArea.AppendChild(valueAxis);

        // Add the plot area to the chart
        chart.AppendChild(plotArea);

        // Add legend
        var legend = new Legend();
        legend.AppendChild(new LegendPosition { Val = LegendPositionValues.Right });

        // Add the legend to the chart
        chart.AppendChild(legend);

        // Add the chart to the chart space
        chartSpace.AppendChild(chart);

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
                new A.Offset { X = new Points(x).AsEmus(), Y = new Points(y).AsEmus() },
                new A.Extents { Cx = new Points(width).AsEmus(), Cy = new Points(height).AsEmus() }),
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
        return shapeIds.Count > 0 ? shapeIds.Max() + 1 : 1;
    }
}