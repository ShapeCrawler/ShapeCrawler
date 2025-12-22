using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents a clustered bar chart.
/// </summary>
internal sealed class ClusteredBarChart(
    ChartPart chartPart,
    IList<List<string>> categories,
    IList<(string Name, double[] Values)> seriesData)
{
    /// <summary>
    ///     Generates the clustered bar chart content.
    /// </summary>
    internal void Generate()
    {
        var chartSpace = CreateChartSpace();
        var chart = chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>() !;
        var plotArea = chart.PlotArea!;
        var barChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.BarChart>() !;

        this.AddSeries(barChart);
        AddAxes(plotArea);

        var legend = new Legend();
        legend.AppendChild(new LegendPosition { Val = LegendPositionValues.Right });
        chart.AppendChild(legend);

        chartPart.ChartSpace = chartSpace;
    }

    private static string ColumnLetter(int columnNumber)
    {
        const int alphabetSize = 26;
        const int asciiOffsetForA = 65;
        var columnLetter = new StringBuilder();

        while (columnNumber > 0)
        {
            var modulo = (columnNumber - 1) % alphabetSize;
            columnLetter.Insert(0, (char)(asciiOffsetForA + modulo));
            columnNumber = (columnNumber - modulo) / alphabetSize;
        }

        return columnLetter.ToString();
    }

    private static ChartSpace CreateChartSpace()
    {
        var chartSpace = new ChartSpace(new EditingLanguage { Val = "en-US" }, new RoundedCorners { Val = false });
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart"); // NOSONAR
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main"); // NOSONAR
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"); // NOSONAR

        var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
        chart.AppendChild(new AutoTitleDeleted { Val = false });

        var plotArea = new PlotArea();
        plotArea.Append(new Layout());

        var barChart = new DocumentFormat.OpenXml.Drawing.Charts.BarChart();
        barChart.AppendChild(new BarDirection { Val = BarDirectionValues.Bar });
        barChart.AppendChild(new BarGrouping { Val = BarGroupingValues.Clustered });
        barChart.AppendChild(new VaryColors { Val = false });

        plotArea.AppendChild(barChart);
        chart.AppendChild(plotArea);
        chartSpace.AppendChild(chart);

        return chartSpace;
    }
    
    private static void AddAxes(PlotArea plotArea)
    {
        const uint axisId1 = 1U;
        const uint axisId2 = 2U;

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
    }

    private void AddSeries(DocumentFormat.OpenXml.Drawing.Charts.BarChart barChart)
    {
        for (uint i = 0; i < seriesData.Count; i++)
        {
            var series = new BarChartSeries(
                new Index { Val = i },
                new Order { Val = i });

            var seriesText = new SeriesText();
            seriesText.AppendChild(new NumericValue { Text = seriesData[(int)i].Name });
            series.AppendChild(seriesText);

            var categoryAxisData = new CategoryAxisData();
            this.AddCategories(categoryAxisData);
            series.AppendChild(categoryAxisData);

            var values = new Values();
            this.AddValues(values, i);
            series.AppendChild(values);

            barChart.AppendChild(series);
        }

        const uint axisId1 = 1U;
        const uint axisId2 = 2U;
        barChart.AppendChild(new AxisId { Val = axisId1 });
        barChart.AppendChild(new AxisId { Val = axisId2 });
    }

    private void AddCategories(CategoryAxisData categoryAxisData)
    {
        bool isMultiLevel = categories.Any(c => c.Count > 1);

        if (isMultiLevel)
        {
            this.AddMultiLevelCategories(categoryAxisData);
        }
        else
        {
            this.AddSingleLevelCategories(categoryAxisData);
        }
    }

    private void AddMultiLevelCategories(CategoryAxisData categoryAxisData)
    {
        var categoriesCount = UInt32Value.FromUInt32((uint)categories.Count);
        var multiLevelStringReference = new MultiLevelStringReference();
        int maxLevel = categories.Max(c => c.Count);
        var endColumnLetter = ColumnLetter(maxLevel);
        multiLevelStringReference.AppendChild(new Formula($"Sheet1!$A$1:${endColumnLetter}${categories.Count}"));
        var multiLevelStringCache = new MultiLevelStringCache();
        multiLevelStringCache.AppendChild(new PointCount { Val = categoriesCount });

        for (int levelIndex = 0; levelIndex < maxLevel; levelIndex++)
        {
            var level = new Level();
            for (int catIndex = 0; catIndex < categories.Count; catIndex++)
            {
                var catList = categories[catIndex];

                // Map the innermost (leaf) category to Level 0 and its parents to higher levels,
                // by reversing the list index so OpenXML multi-level categories are leaf-first.
                int listIndex = (catList.Count - 1) - levelIndex;
                if (listIndex < 0)
                {
                    continue;
                }

                var value = catList[listIndex];
                var point = new StringPoint { Index = (uint)catIndex };
                point.AppendChild(new NumericValue(value));
                level.AppendChild(point);
            }

            multiLevelStringCache.AppendChild(level);
        }

        multiLevelStringReference.AppendChild(multiLevelStringCache);
        categoryAxisData.AppendChild(multiLevelStringReference);
    }

    private void AddSingleLevelCategories(CategoryAxisData categoryAxisData)
    {
        var categoriesCount = UInt32Value.FromUInt32((uint)categories.Count);
        var stringLiteral = new StringLiteral(new PointCount { Val = categoriesCount });
        for (uint j = 0; j < categories.Count; j++)
        {
            var point = new StringPoint { Index = j };
            point.AppendChild(new NumericValue(categories[(int)j][0]));
            stringLiteral.AppendChild(point);
        }

        categoryAxisData.AppendChild(stringLiteral);
    }

    private void AddValues(Values values, uint seriesIndex)
    {
        var categoriesCount = UInt32Value.FromUInt32((uint)categories.Count);
        var numberLiteral = new NumberLiteral(new FormatCode("General"), new PointCount { Val = categoriesCount });

        for (uint j = 0; j < seriesData[(int)seriesIndex].Values.Length; j++)
        {
            var point = new NumericPoint { Index = j };
            point.AppendChild(new NumericValue(seriesData[(int)seriesIndex].Values[j].ToString(CultureInfo.InvariantCulture)));
            numberLiteral.AppendChild(point);
        }

        values.AppendChild(numberLiteral);
    }
}
