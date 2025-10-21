using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Slides;

/// <summary>
///     Represents the content of a bar chart.
/// </summary>
internal sealed class BarChartContent(
    ChartPart chartPart,
    Dictionary<string, double> categoryValues,
    string seriesName)
{
    /// <summary>
    ///     Generates the bar chart content.
    /// </summary>
    public void Generate()
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
}
