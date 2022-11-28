using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Exceptions;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable PossibleMultipleEnumeration
// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a chart series.
/// </summary>
public interface ISeries
{
    /// <summary>
    ///     Gets series name.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets chart type.
    /// </summary>
    SCChartType Type { get; }

    /// <summary>
    ///     Gets collection of chart points.
    /// </summary>
    IChartPointCollection Points { get; }

    /// <summary>
    ///     Gets a value indicating whether chart has name.
    /// </summary>
    bool HasName { get; }
}

internal class Series : ISeries
{
    private readonly Lazy<string?> name;
    private readonly OpenXmlElement seriesXmlElement;
    private readonly SCChart parentChart;

    internal Series(SCChart parentChart, OpenXmlElement seriesXmlElement, SCChartType seriesChartType)
    {
        this.parentChart = parentChart;
        this.seriesXmlElement = seriesXmlElement;
        this.name = new Lazy<string?>(this.GetNameOrDefault);
        this.Type = seriesChartType;
    }

    public SCChartType Type { get; }

    public IChartPointCollection Points
    {
        get
        {
            ErrorHandler.Execute(() => ChartPointCollection.Create(this.parentChart, this.seriesXmlElement), out var result);
            return result;
        }
    }

    public bool HasName => this.name.Value != null;

    public string Name
    {
        get
        {
            if (this.name.Value == null)
            {
                throw new NotSupportedException(ExceptionMessages.SeriesHasNotName);
            }

            return this.name.Value;
        }
    }

    private string? GetNameOrDefault()
    {
        var cStringReference = this.seriesXmlElement.GetFirstChild<C.SeriesText>()?.StringReference;
        if (cStringReference == null)
        {
            return null;
        }

        return ChartReferencesParser.GetSingleString(cStringReference, this.parentChart);
    }
}