using SlideXML.Enums;

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a chart.
    /// </summary>
    public interface IChart
    {
        ChartType Type { get; }

        string Title { get; }
    }
}