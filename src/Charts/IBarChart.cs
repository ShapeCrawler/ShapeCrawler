#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a bar chart.
/// </summary>
public interface IBarChart : IChart
{
}

/// <summary>
///     Represents a column chart.
/// </summary>
public interface IColumnChart : IChart
{
}

/// <summary>
///     Represents a line chart.
/// </summary>
public interface ILineChart : IChart
{
}

/// <summary>
///     Represents a pie chart.
/// </summary>
public interface IPieChart : IChart
{
}

/// <summary>
///     Represents a scatter chart.
/// </summary>
public interface IScatterChart : IChart
{
}

/// <summary>
///     Represents a bubble chart.
/// </summary>
public interface IBubbleChart : IChart
{
}

/// <summary>
///     Represents an area chart.
/// </summary>
public interface IAreaChart : IChart
{
}