using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Presentations;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Slides;

internal sealed class ChartCollection(SlidePart slidePart)
{
    public void AddPieChart(
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
        GenerateBarChartContent(chartPart, categoryValues, seriesName);
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
        GenerateScatterChartContent(chartPart, pointValues, seriesName);
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
        GenerateStackedColumnChartContent(chartPart, categoryValues, seriesNames);
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Stacked Column Chart");
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

        pieChart.Append(
            new DataLabels(
                new ShowLegendKey { Val = false },
                new ShowValue { Val = true },
                new ShowCategoryName { Val = false },
                new ShowSeriesName { Val = false },
                new ShowPercent { Val = false },
                new ShowBubbleSize { Val = false },
                new ShowLeaderLines { Val = true }
            )
        );
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

    private static void GenerateScatterChartContent(
        ChartPart chartPart,
        Dictionary<double, double> pointValues,
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
        var series = new ScatterChartSeries(new Index { Val = 0 }, new Order { Val = 0 });

        // Add series name
        var seriesText = new SeriesText();
        seriesText.AppendChild(new NumericValue { Text = seriesName });
        series.AppendChild(seriesText);

        // --- X Values ---
        var pointsCount = UInt32Value.FromUInt32((uint)pointValues.Count);

        var xValues = new XValues();
        var xNumberLiteral = new NumberLiteral
        {
            FormatCode = new FormatCode("General"), PointCount = new PointCount { Val = pointsCount }
        };

        uint index = 0;
        foreach (var point in pointValues)
        {
            var numericPoint = new NumericPoint { Index = index };
            numericPoint.AppendChild(new NumericValue(point.Key.ToString()));
            xNumberLiteral.AppendChild(numericPoint);
            index++;
        }

        xValues.AppendChild(xNumberLiteral);
        series.AppendChild(xValues);

        // --- Y Values ---
        var yValues = new YValues();
        var yNumberLiteral = new NumberLiteral
        {
            FormatCode = new FormatCode("General"), PointCount = new PointCount { Val = pointsCount }
        };

        index = 0;
        foreach (var point in pointValues)
        {
            var numericPoint = new NumericPoint { Index = index };
            numericPoint.AppendChild(new NumericValue(point.Value.ToString()));
            yNumberLiteral.AppendChild(numericPoint);
            index++;
        }

        yValues.AppendChild(yNumberLiteral);
        series.AppendChild(yValues);

        // Define the axes IDs
        const uint axisId1 = 1U;
        const uint axisId2 = 2U;

        // Create the scatter chart and add it to the plot area
        var plotArea = new PlotArea(
            new Layout(),
            new ScatterChart(
                new ScatterStyle { Val = ScatterStyleValues.LineMarker },
                new VaryColors { Val = false },
                series,
                new AxisId { Val = axisId1 },
                new AxisId { Val = axisId2 }));

        // Add X axis (Value Axis for X)
        var xAxis = new ValueAxis();
        xAxis.AppendChild(new AxisId { Val = axisId1 });

        var scaling = new Scaling();
        scaling.AppendChild(new Orientation { Val = OrientationValues.MinMax });
        xAxis.AppendChild(scaling);

        xAxis.AppendChild(new Delete { Val = false });
        xAxis.AppendChild(new AxisPosition { Val = AxisPositionValues.Bottom });
        xAxis.AppendChild(new MajorGridlines());
        xAxis.AppendChild(new NumberingFormat { FormatCode = "General", SourceLinked = true });
        xAxis.AppendChild(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        xAxis.AppendChild(new CrossingAxis { Val = axisId2 });
        xAxis.AppendChild(new Crosses { Val = CrossesValues.AutoZero });
        xAxis.AppendChild(new CrossBetween { Val = CrossBetweenValues.Between });

        // Add the X axis to the plot area
        plotArea.AppendChild(xAxis);

        // Add Y axis (Value Axis for Y)
        var yAxis = new ValueAxis();
        yAxis.AppendChild(new AxisId { Val = axisId2 });

        scaling = new Scaling();
        scaling.AppendChild(new Orientation { Val = OrientationValues.MinMax });
        yAxis.AppendChild(scaling);

        yAxis.AppendChild(new Delete { Val = false });
        yAxis.AppendChild(new AxisPosition { Val = AxisPositionValues.Left });
        yAxis.AppendChild(new MajorGridlines());
        yAxis.AppendChild(new NumberingFormat { FormatCode = "General", SourceLinked = true });
        yAxis.AppendChild(new TickLabelPosition { Val = TickLabelPositionValues.NextTo });
        yAxis.AppendChild(new CrossingAxis { Val = axisId1 });
        yAxis.AppendChild(new Crosses { Val = CrossesValues.AutoZero });
        yAxis.AppendChild(new CrossBetween { Val = CrossBetweenValues.Between });

        // Add the Y axis to the plot area
        plotArea.AppendChild(yAxis);

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

    private static void GenerateStackedColumnChartContent(
        ChartPart chartPart,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames)
    {
        // Create the ChartSpace element
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var chart = new Chart();
        chart.AppendChild(new AutoTitleDeleted { Val = false });

        // Create a plot area
        var plotArea = new PlotArea();
        plotArea.AppendChild(new Layout());

        // Create the bar chart
        var barChart = new BarChart(
            new BarDirection { Val = BarDirectionValues.Column },
            new BarGrouping { Val = BarGroupingValues.Stacked },
            new VaryColors { Val = false });

        // Create a list to store the categories for reuse
        var categories = categoryValues.Keys.ToList();

        // Add series
        for (int i = 0; i < seriesNames.Count; i++)
        {
            var series = new BarChartSeries(
                new Index { Val = (uint)i },
                new Order { Val = (uint)i });

            // Add series name - use simple text value instead of StringLiteral
            var seriesText = new SeriesText();
            seriesText.AppendChild(new NumericValue(seriesNames[i]));
            series.AppendChild(seriesText);

            // Add category axis data
            var categoryAxisData = new CategoryAxisData();
            var stringReference = new StringReference();
            stringReference.AppendChild(new Formula($"Sheet1!$A$2:$A${categories.Count + 1}"));
            var stringCache = new StringCache();
            stringCache.AppendChild(new PointCount { Val = (uint)categories.Count });

            for (uint j = 0; j < categories.Count; j++)
            {
                var point = new StringPoint { Index = j };
                point.AppendChild(new NumericValue(categories[(int)j]));
                stringCache.AppendChild(point);
            }

            stringReference.AppendChild(stringCache);
            categoryAxisData.AppendChild(stringReference);
            series.AppendChild(categoryAxisData);

            // Add values
            var values = new Values();
            var numberReference = new NumberReference();

            // Important to use a valid Excel range formula here
            numberReference.AppendChild(new Formula($"Sheet1!$B$2:$B${categories.Count + 1}"));
            var numberCache = new NumberingCache();
            numberCache.AppendChild(new FormatCode("General"));
            numberCache.AppendChild(new PointCount { Val = (uint)categories.Count });

            for (uint j = 0; j < categories.Count; j++)
            {
                var point = new NumericPoint { Index = j };
                point.AppendChild(new NumericValue(categoryValues[categories[(int)j]][i]
                    .ToString(CultureInfo.InvariantCulture)));
                numberCache.AppendChild(point);
            }

            numberReference.AppendChild(numberCache);
            values.AppendChild(numberReference);
            series.AppendChild(values);

            barChart.AppendChild(series);
        }

        // Add axis IDs
        const uint axisId1 = 1U;
        const uint axisId2 = 2U;
        barChart.AppendChild(new AxisId { Val = axisId1 });
        barChart.AppendChild(new AxisId { Val = axisId2 });

        // Add the bar chart to the plot area
        plotArea.AppendChild(barChart);

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
        legend.AppendChild(new Layout());
        chart.AppendChild(legend);

        // Add the chart to the chart space
        chartSpace.AppendChild(chart);

        // Create embedded package part for data
        var excelPackagePart = chartPart.AddNewPart<EmbeddedPackagePart>(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Microsoft_Excel_Sheet1");
        using (var stream = excelPackagePart.GetStream(FileMode.Create, FileAccess.Write))
        {
            // Just reserve a placeholder - the actual Excel file isn't needed
            // but the package part must exist for the chart to work correctly
            stream.Write(new byte[10], 0, 10);
        }

        // Save the chart part
        chartPart.ChartSpace = chartSpace;
    }

    private void InsertChartGraphicFrame(ChartPart chartPart, int x, int y, int width, int height)
    {
        var graphicFrame = new GraphicFrame
        {
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