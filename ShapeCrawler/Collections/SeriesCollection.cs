using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a chart series collection.
    /// </summary>
    public class SeriesCollection : LibraryCollection<Series>
    {
        internal SeriesCollection(List<Series> seriesList)
        {
            CollectionItems = seriesList.ToList();
        }

        internal static SeriesCollection Create(
            IEnumerable<OpenXmlElement> cXCharts,
            ChartPart chartPart,
            ChartReferencesParser chartRefParser)
        {
            var seriesList = new List<Series>();
            foreach (OpenXmlElement cXChart in cXCharts)
            {
                Enum.TryParse(cXChart.LocalName, true, out ChartType chartType);
                var nextSdkChartSeriesCollection = cXChart.ChildElements
                    .Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
                foreach (var sdkSeries in nextSdkChartSeriesCollection)
                {
                    var series = new Series(chartType, sdkSeries, chartPart, chartRefParser);
                    seriesList.Add(series);
                }
            }

            return new SeriesCollection(seriesList);
        }
    }
}