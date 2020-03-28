using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Models.SlideComponents.Chart;
using SlideDotNet.Validation;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Collections
{
    /// <summary>
    /// Represents a collection of the chart series.
    /// </summary>
    public class SeriesCollection : LibraryCollection<Series>
    {
        #region Constructors

        /// <summary>
        /// Initializes a new collection of the chart series.
        /// </summary>
        public SeriesCollection(IEnumerable<OpenXmlElement> sdkCharts, ChartPart sdkChartPart)
        {
            Check.NotEmpty(sdkCharts, nameof(sdkCharts));

            var tempSeriesCollection = new LinkedList<Series>();
            foreach (var nextSdkChart in sdkCharts)
            {
                Enum.TryParse(nextSdkChart.LocalName, true, out ChartType chartType);
                var nextSdkChartSeriesCollection = nextSdkChart.ChildElements.Where(e => e.LocalName.Equals("ser"));
                foreach (var sdkSeries in nextSdkChartSeriesCollection)
                {
                    var series = new Series(chartType, sdkSeries, sdkChartPart);
                    tempSeriesCollection.AddLast(series);
                }
            }

            CollectionItems = new List<Series>(tempSeriesCollection);
        }

        #endregion Constructors
    }
}