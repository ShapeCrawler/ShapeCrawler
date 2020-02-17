using SlideDotNet.Enums;

namespace SlideDotNet.Models.SlideComponents.Chart
{
    /// <summary>
    /// Represents a chart.
    /// </summary>
    public interface IChart
    {
        ChartType Type { get; }

        string Title { get; }

        bool HasTitle { get; }
    }
}