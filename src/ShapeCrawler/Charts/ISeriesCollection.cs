using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collection of chart series.
/// </summary>
public interface ISeriesCollection : IReadOnlyCollection<ISeries>
{
    /// <summary>
    ///     Gets the series at the specified index.
    /// </summary>
    ISeries this[int index] { get; }

    /// <summary>
    ///     Removes the series at the specified index.
    /// </summary>
    void RemoveAt(int index);
}

internal sealed class SCSeriesCollection : ISeriesCollection
{
    private readonly List<ISeries> seriesList;

    internal SCSeriesCollection(List<ISeries> seriesList)
    {
        this.seriesList = seriesList;
    }
    
    public int Count => this.seriesList.Count;
    
    public ISeries this[int index] => this.seriesList[index];
    
    public void RemoveAt(int index)
    {
        var seriesCore = (SCSeries)this.seriesList[index];
        seriesCore.CSer.Remove();
        
        this.seriesList.RemoveAt(index);
    }

    public IEnumerator<ISeries> GetEnumerator()
    {
        return this.seriesList.GetEnumerator();
    }
    
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
    
    internal static SCSeriesCollection Create(SCChart slideChart, IEnumerable<OpenXmlElement> cXCharts)
    {
        var seriesList = new List<ISeries>();
        foreach (var cXChart in cXCharts)
        {
            Enum.TryParse(cXChart.LocalName, true, out SCChartType seriesChartType);
            var cSerList = cXChart.ChildElements.Where(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
            seriesList.AddRange(cSerList.Select(cSer => new SCSeries(slideChart, cSer, seriesChartType)));
        }

        return new SCSeriesCollection(seriesList);
    }
}