using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents the content of a stacked column chart.
/// </summary>
internal sealed class StackedColumnChart(
    ChartPart chartPart,
    IDictionary<string, IList<double>> categoryValues,
    IList<string> seriesNames)
{
    /// <summary>
    ///     Generates the stacked column chart content.
    /// </summary>
    public void Generate() // TODO: move to constructor
    {
        // Create the ChartSpace element
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        chart.AppendChild(new AutoTitleDeleted { Val = false });

        // Create a plot area
        var plotArea = new PlotArea();
        plotArea.AppendChild(new Layout());

        // Create the bar chart
        var barChart = new DocumentFormat.OpenXml.Drawing.Charts.BarChart(
            new BarDirection { Val = BarDirectionValues.Column },
            new BarGrouping { Val = BarGroupingValues.Stacked },
            new VaryColors { Val = false });

        // Create a list to store the categories for reuse
        var categories = categoryValues.Keys.ToList();

        // Add series
        for (var i = 0; i < seriesNames.Count; i++)
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
}