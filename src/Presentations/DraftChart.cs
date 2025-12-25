using System.Collections.Generic;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft chart builder.
/// </summary>
public sealed class DraftChart
{
    internal string ChartName { get; private set; } = "Chart";

    internal int ChartX { get; private set; } = 100;

    internal int ChartY { get; private set; } = 100;

    internal int ChartWidth { get; private set; } = 400;

    internal int ChartHeight { get; private set; } = 300;

    internal List<List<string>> CategoryNames { get; } = [];

    internal List<SeriesData> SeriesDataList { get; } = [];

    /// <summary>
    ///     Sets chart name.
    /// </summary>
    public DraftChart Name(string name)
    {
        this.ChartName = name;
        return this;
    }

    /// <summary>
    ///     Sets X-position.
    /// </summary>
    public DraftChart X(int x)
    {
        this.ChartX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position.
    /// </summary>
    public DraftChart Y(int y)
    {
        this.ChartY = y;
        return this;
    }

    /// <summary>
    ///     Sets width.
    /// </summary>
    public DraftChart Width(int width)
    {
        this.ChartWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height.
    /// </summary>
    public DraftChart Height(int height)
    {
        this.ChartHeight = height;
        return this;
    }

    /// <summary>
    ///     Adds categories to the chart.
    /// </summary>
    public DraftChart Categories(params string[] categories)
    {
        if (categories is null)
        {
            throw new SCException($"{nameof(categories)} cannot be null.");
        }
        
        foreach (var category in categories)
        {
            this.CategoryNames.Add([category]);
        }

        return this;
    }

    /// <summary>
    ///     Adds multi-level categories to the chart.
    /// </summary>
    public DraftChart Categories(params (string Main, string Sub)[] categories)
    {
        foreach (var (main, sub) in categories)
        {
            this.CategoryNames.Add([main, sub]);
        }

        return this;
    }

    /// <summary>
    ///     Adds a series with values to the chart.
    /// </summary>
    public DraftChart Series(string seriesName, params double[] values)
    {
        this.SeriesDataList.Add(new SeriesData(seriesName, values));
        return this;
    }

    /// <summary>
    ///     Represents series data for a chart.
    /// </summary>
    public sealed class SeriesData(string name, double[] values)
    {
        internal string Name { get; } = name;

        internal double[] Values { get; } = values;
    }
}