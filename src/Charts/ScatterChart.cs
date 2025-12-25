using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents the content of a scatter chart.
/// </summary>
internal sealed class ScatterChart(
    ChartPart chartPart,
    Dictionary<double, double> pointValues,
    string seriesName)
{
    /// <summary>
    ///     Generates the scatter chart content.
    /// </summary>
    public void Generate()
    {
        // Create the ChartSpace element
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
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
            FormatCode = new FormatCode("General"),
            PointCount = new PointCount { Val = pointsCount }
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
            FormatCode = new FormatCode("General"),
            PointCount = new PointCount { Val = pointsCount }
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
            new DocumentFormat.OpenXml.Drawing.Charts.ScatterChart(
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
}