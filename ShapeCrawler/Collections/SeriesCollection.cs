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
    ///     Represents chart series collection.
    /// </summary>
    public class SeriesCollection : LibraryCollection<Series>
    {
        #region Constructors

        /// <summary>
        ///     Initializes a new collection of the chart series.
        /// </summary>
        internal SeriesCollection(IEnumerable<OpenXmlElement> sdkCharts, ChartPart sdkChartPart,
            ChartRefParser chartRefParser)
        {
            var tempSeriesCollection = new LinkedList<Series>(); //TODO: make weak reference
            foreach (var nextSdkChart in sdkCharts)
            {
                Enum.TryParse(nextSdkChart.LocalName, true, out ChartType chartType);
                var nextSdkChartSeriesCollection = nextSdkChart.ChildElements
                    .Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
                foreach (var sdkSeries in nextSdkChartSeriesCollection)
                {
                    var series = new Series(chartType, sdkSeries, sdkChartPart, chartRefParser);
                    tempSeriesCollection.AddLast(series);
                }
            }

            CollectionItems = new List<Series>(tempSeriesCollection);
        }

        #endregion Constructors
    }
}