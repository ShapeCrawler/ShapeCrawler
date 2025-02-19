using System.Collections.Generic;
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
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        GeneratePieChartContent(chartPart, categoryValues, seriesName);
        this.InsertChartGraphicFrame(chartPart, x, y, width, height);
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

    private void InsertChartGraphicFrame(ChartPart chartPart, int x, int y, int width, int height)
    {
        // Create a new GraphicFrame
        var graphicFrame = new GraphicFrame
        {
            // Give it an ID and name
            NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = 2U, Name = "Pie Chart" },
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
}