using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents a numeric data collection in a chart series.
/// </summary>
/// <param name="chartPart">The chart part owning the series.</param>
/// <param name="series">The series Open XML element.</param>
/// <param name="valuesName">The numeric data element name, or null for the series values.</param>
internal sealed class ChartPoints(ChartPart chartPart, OpenXmlElement series, string? valuesName = null)
    : IChartPointCollection
{
    /// <inheritdoc />
    public int Count => this.Points().Count;

    private OpenXmlCompositeElement? Values => valuesName == null
        ? (OpenXmlCompositeElement?)series.GetFirstChild<C.Values>() ?? series.GetFirstChild<C.YValues>()
        : series.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(element => element.LocalName == valuesName);

    /// <inheritdoc />
    public IChartPoint this[int index] => this.Points()[index];

    /// <inheritdoc />
    public void Add(double value) => new ChartPointAddition(chartPart, this.Values).Add(value);

    /// <inheritdoc />
    public IEnumerator<IChartPoint> GetEnumerator() => this.Points().GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<ChartPoint> Points()
    {
        var values = this.Values;
        return new ChartPointData(chartPart).Create(
            values?.GetFirstChild<C.NumberReference>(), values?.GetFirstChild<C.NumberLiteral>());
    }
}