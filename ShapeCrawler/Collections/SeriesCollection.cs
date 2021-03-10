using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of series.
    /// </summary>
    internal class SeriesCollection : LibraryCollection<Series>, ISeriesCollection
    {
        internal SeriesCollection(List<Series> seriesList)
        {
            CollectionItems = seriesList;
        }

        internal static SeriesCollection Create(
            SlideChart slideChart, 
            IEnumerable<OpenXmlElement> cXCharts,
            ChartReferencesParser chartRefParser)
        {
            var seriesList = new List<Series>();
            foreach (OpenXmlElement cXChart in cXCharts)
            {
                Enum.TryParse(cXChart.LocalName, true, out ChartType chartType); //TODO: use Parse instead of TryParse
                var nextSdkChartSeriesCollection = cXChart.ChildElements
                    .Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
                foreach (OpenXmlElement seriesXmlElement in nextSdkChartSeriesCollection)
                {
                    var series = new Series(slideChart, chartType, seriesXmlElement, chartRefParser);
                    seriesList.Add(series);
                }
            }

            return new SeriesCollection(seriesList);
        }
    }
}