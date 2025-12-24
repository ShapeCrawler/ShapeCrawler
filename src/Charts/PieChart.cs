using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents the content of a pie chart.
/// </summary>
internal sealed class PieChart(
    ChartPart chartPart,
    Dictionary<string, double> categoryValues,
    string seriesName)
{
    /// <summary>
    ///     Generates the pie chart content.
    /// </summary>
    public void Generate()
    {
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        chart.AppendChild(new AutoTitleDeleted { Val = false });

        var plotArea = new PlotArea();
        plotArea.Append(new Layout());

        var pieChart = new DocumentFormat.OpenXml.Drawing.Charts.PieChart();
        pieChart.Append(new VaryColors { Val = true });

        var seriesText = new SeriesText();
        var stringReference = new StringReference();
        stringReference.Append(new Formula("Sheet1!$A$1"));
        var stringCache = new StringCache();
        stringCache.Append(new PointCount { Val = 1 });
        stringCache.Append(new StringPoint(new NumericValue(seriesName)) { Index = 0 });
        stringReference.Append(stringCache);
        seriesText.Append(stringReference);

        var series = new PieChartSeries(
            new Index { Val = 0 },
            new Order { Val = 0 },
            seriesText);

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
}