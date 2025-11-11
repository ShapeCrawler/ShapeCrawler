using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents a clustered bar chart generator.
/// </summary>
internal sealed class ClusteredBarChart(
    ChartPart chartPart,
    List<string> categories,
    List<(string Name, double[] Values)> seriesData)
{
    /// <summary>
    ///     Generates the clustered bar chart content.
    /// </summary>
    public void Generate()
    {
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        chart.AppendChild(new AutoTitleDeleted { Val = false });

        var plotArea = new PlotArea();
        plotArea.Append(new Layout());

        var barChart = new DocumentFormat.OpenXml.Drawing.Charts.BarChart();
        barChart.AppendChild(new BarDirection { Val = BarDirectionValues.Bar });
        barChart.AppendChild(new BarGrouping { Val = BarGroupingValues.Clustered });
        barChart.AppendChild(new VaryColors { Val = false });

        var categoriesCount = UInt32Value.FromUInt32((uint)categories.Count);

        for (uint i = 0; i < seriesData.Count; i++)
        {
            var series = new BarChartSeries(
                new Index { Val = i },
                new Order { Val = i });

            var seriesText = new SeriesText();
            seriesText.AppendChild(new NumericValue { Text = seriesData[(int)i].Name });
            series.AppendChild(seriesText);

            var categoryAxisData = new CategoryAxisData();
            var stringLiteral = new StringLiteral(new PointCount { Val = categoriesCount });

            for (uint j = 0; j < categories.Count; j++)
            {
                var point = new StringPoint { Index = j };
                point.AppendChild(new NumericValue(categories[(int)j]));
                stringLiteral.AppendChild(point);
            }

            categoryAxisData.AppendChild(stringLiteral);
            series.AppendChild(categoryAxisData);

            var values = new Values();
            var numberLiteral = new NumberLiteral(new FormatCode("General"), new PointCount { Val = categoriesCount });

            for (uint j = 0; j < seriesData[(int)i].Values.Length; j++)
            {
                var point = new NumericPoint { Index = j };
                point.AppendChild(new NumericValue(seriesData[(int)i].Values[j].ToString(CultureInfo.InvariantCulture)));
                numberLiteral.AppendChild(point);
            }

            values.AppendChild(numberLiteral);
            series.AppendChild(values);

            barChart.AppendChild(series);
        }

        const uint axisId1 = 1U;
        const uint axisId2 = 2U;
        barChart.AppendChild(new AxisId { Val = axisId1 });
        barChart.AppendChild(new AxisId { Val = axisId2 });

        plotArea.AppendChild(barChart);

        var categoryAxis = new CategoryAxis();
        categoryAxis.AppendChild(new AxisId { Val = axisId1 });

        var scaling = new Scaling();
        scaling.AppendChild(new Orientation { Val = OrientationValues.MinMax });
        categoryAxis.AppendChild(scaling);

        categoryAxis.AppendChild(new Delete { Val = false });
        categoryAxis.AppendChild(new AxisPosition { Val = AxisPositionValues.Bottom });
        categoryAxis.AppendChild(new CrossingAxis { Val = axisId2 });

        plotArea.AppendChild(categoryAxis);

        var valueAxis = new ValueAxis();
        valueAxis.AppendChild(new AxisId { Val = axisId2 });

        scaling = new Scaling();
        scaling.AppendChild(new Orientation { Val = OrientationValues.MinMax });
        valueAxis.AppendChild(scaling);

        valueAxis.AppendChild(new Delete { Val = false });
        valueAxis.AppendChild(new AxisPosition { Val = AxisPositionValues.Left });
        valueAxis.AppendChild(new CrossingAxis { Val = axisId1 });

        plotArea.AppendChild(valueAxis);

        chart.AppendChild(plotArea);

        var legend = new Legend();
        legend.AppendChild(new LegendPosition { Val = LegendPositionValues.Right });

        chart.AppendChild(legend);

        chartSpace.AppendChild(chart);

        chartPart.ChartSpace = chartSpace;
    }
}
