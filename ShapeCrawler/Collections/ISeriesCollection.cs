using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of chart series.
    /// </summary>
    public interface ISeriesCollection : IEnumerable<ISeries>
    {
        /// <summary>
        ///     Gets the number of series items in the collection.
        /// </summary>
        int Count { get; }

        /// <summary>
        ///     Gets the element at the specified index.
        /// </summary>
        ISeries this[int index] { get; }
    }

    internal class SeriesCollection : LibraryCollection<ISeries>, ISeriesCollection
    {
        internal SeriesCollection(List<ISeries> seriesList)
        {
            this.CollectionItems = seriesList;
        }

        internal static SeriesCollection Create(SCChart slideChart, IEnumerable<OpenXmlElement> cXCharts)
        {
            var seriesList = new List<ISeries>();
            foreach (OpenXmlElement cXChart in cXCharts)
            {
                Enum.TryParse(cXChart.LocalName, true, out ChartType seriesChartType); // TODO: use Parse instead of TryParse
                IEnumerable<OpenXmlElement> nextSdkChartSeriesCollection = cXChart.ChildElements
                    .Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
                seriesList.AddRange(nextSdkChartSeriesCollection.Select(seriesXmlElement =>
                    new Series(slideChart, seriesXmlElement, seriesChartType)));
            }

            return new SeriesCollection(seriesList);
        }
    }
}