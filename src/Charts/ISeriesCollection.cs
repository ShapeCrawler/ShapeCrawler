﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of chart series.
/// </summary>
public interface ISeriesCollection : IReadOnlyList<ISeries>
{
    /// <summary>
    ///     Removes the series at the specified index.
    /// </summary>
    void RemoveAt(int index);
}

internal sealed class SeriesCollection : ISeriesCollection
{
    private readonly ChartPart chartPart;
    private readonly IEnumerable<OpenXmlElement> cXCharts;

    internal SeriesCollection(ChartPart chartPart, IEnumerable<OpenXmlElement> cXCharts)
    {
        this.chartPart = chartPart;
        this.cXCharts = cXCharts;
    }
    
    public int Count => this.SeriesListCore().Count;
    
    public ISeries this[int index] => this.SeriesListCore()[index];
    
    public void RemoveAt(int index) => this.CSerList()[index].Remove();

    public IEnumerator<ISeries> GetEnumerator() => this.SeriesListCore().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
    
    private List<ISeries> SeriesListCore()
    {
        var seriesList = new List<ISeries>();
        foreach (var cXChart in this.cXCharts)
        {
            Enum.TryParse(cXChart.LocalName, true, out ChartType seriesChartType);
            var cSerList = cXChart.ChildElements.Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
            seriesList.AddRange(cSerList.Select(cSer => new Series(this.chartPart, cSer, seriesChartType)));
        }

        return seriesList;
    }

    private List<OpenXmlElement> CSerList()
    {
        var cSerList = new List<OpenXmlElement>();
        foreach (var cXChart in this.cXCharts)
        {
            var chartCSerList = cXChart.ChildElements.Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
            cSerList.AddRange(chartCSerList);
        }

        return cSerList;
    }
}