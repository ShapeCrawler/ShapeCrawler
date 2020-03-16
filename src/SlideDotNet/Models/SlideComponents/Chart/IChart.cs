using System.Collections.Generic;
using SlideDotNet.Collections;
using SlideDotNet.Enums;

namespace SlideDotNet.Models.SlideComponents.Chart
{
    /// <summary>
    /// Represents a chart.
    /// </summary>
    public interface IChart
    {
        /// <summary>
        /// Returns type of the chart.
        /// </summary>
        ChartType Type { get; }

        /// <summary>
        /// Returns the chart title. Returns null if chart has not a title.
        /// </summary>
        string Title { get; }

        /// <summary>
        /// Indicates whether chart has a title.
        /// </summary>
        bool HasTitle { get; }

        /// <summary>
        /// Returns collection of the chart series.
        /// </summary>
        SeriesCollection SeriesCollection { get; }

        bool HasCategories { get; }

        /// <summary>
        /// Returns collection of the chart category.
        /// </summary>
        CategoryCollection Categories { get; }
    }
}