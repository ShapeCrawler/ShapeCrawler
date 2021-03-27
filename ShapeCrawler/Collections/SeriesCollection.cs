using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;

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

        internal static SeriesCollection Create(SlideChart slideChart, IEnumerable<OpenXmlElement> cXCharts)
        {
            var seriesList = new List<Series>();
            foreach (OpenXmlElement cXChart in cXCharts)
            {
                Enum.TryParse(cXChart.LocalName, true, out ChartType chartType); //TODO: use Parse instead of TryParse
                IEnumerable<OpenXmlElement> nextSdkChartSeriesCollection = cXChart.ChildElements
                    .Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
                seriesList.AddRange(nextSdkChartSeriesCollection.Select(seriesXmlElement => new Series(slideChart, chartType, seriesXmlElement)));
            }

            return new SeriesCollection(seriesList);
        }
    }
}