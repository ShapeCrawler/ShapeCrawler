using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collection of chart series.
/// </summary>
public interface ISeriesList : IReadOnlyList<ISeries>
{
    /// <summary>
    ///     Removes the series at the specified index.
    /// </summary>
    void RemoveAt(int index);
}

internal sealed class SeriesList : ISeriesList
{
    private readonly ChartPart sdkChartPart;
    private readonly IEnumerable<OpenXmlElement> cXCharts;

    internal SeriesList(ChartPart sdkChartPart, IEnumerable<OpenXmlElement> cXCharts)
    {
        this.sdkChartPart = sdkChartPart;
        this.cXCharts = cXCharts;
    }

    public int Count => this.SeriesListCore().Count;

    private List<ISeries> SeriesListCore()
    {
        var seriesList = new List<ISeries>();
        foreach (var cXChart in cXCharts)
        {
            Enum.TryParse(cXChart.LocalName, true, out SCChartType seriesChartType);
            seriesList.AddRange(this.CSerList().Select(cSer => new Series(this.sdkChartPart, cSer, seriesChartType)));
        }

        return seriesList;
    }

    private List<OpenXmlElement> CSerList()
    {
        var cSerList = new List<OpenXmlElement>();
        foreach (var cXChart in cXCharts)
        {
            var chartCSerList = cXChart.ChildElements.Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
            cSerList.AddRange(chartCSerList);
        }

        return cSerList;
    }

    public ISeries this[int index] => this.SeriesListCore()[index];

    public void RemoveAt(int index) => this.CSerList()[index].Remove();

    public IEnumerator<ISeries> GetEnumerator()
    {
        return this.SeriesListCore().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}