using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

/// <summary>
///     Adds a series to a supported template chart.
/// </summary>
internal sealed class ChartSeriesAddition(ChartPart chartPart, IEnumerable<OpenXmlElement> chartElements)
{
    /// <summary>
    ///     Adds a series using the chart's existing categories and template formatting.
    /// </summary>
    /// <param name="name">The name of the series.</param>
    /// <param name="values">One value for each existing category or template X value.</param>
    internal void Add(string name, IReadOnlyList<double> values)
    {
        var charts = chartElements.ToList();
        if (charts.Count != 1 || !this.IsSupportedChart(charts[0].LocalName))
        {
            throw new SCException("Adding a series is supported only for non-combination bar, line, pie/doughnut, scatter, and area charts.");
        }

        var chart = charts[0];
        var isScatterChart = chart.LocalName.Equals("scatterChart", StringComparison.Ordinal);
        var seriesList = chart.ChildElements.Where(element => element.LocalName.Equals("ser", StringComparison.Ordinal)).ToList();
        if (seriesList.Count == 0)
        {
            throw new SCException("Cannot add a series without an existing template series.");
        }

        var template = seriesList[0];
        var templateSeriesText = template.GetFirstChild<C.SeriesText>();
        var nameFormula = templateSeriesText?.StringReference?.Formula?.Text;
        var valueElementName = isScatterChart ? "yVal" : "val";
        var templateValues = this.NumberData(template, valueElementName);
        var valuesFormula = templateValues.GetFirstChild<C.NumberReference>()?.Formula?.Text;
        this.ValidateValueCount(template, isScatterChart, values.Count);

        var hasNameFormula = !string.IsNullOrWhiteSpace(nameFormula);
        var hasValuesFormula = !string.IsNullOrWhiteSpace(valuesFormula);
        if ((hasValuesFormula && !hasNameFormula) || (hasValuesFormula && chartPart.EmbeddedPackagePart is null))
        {
            throw new SCException("Updating a formula-backed series requires both formulas and an embedded Excel workbook.");
        }

        string? newNameFormula = null;
        string? newValuesFormula = null;
        Workbook? workbook = null;
        if (hasValuesFormula)
        {
            workbook = new Workbook(chartPart.EmbeddedPackagePart!);
            var newColumn = this.NextSeriesColumn(seriesList, valueElementName, workbook);
            newNameFormula = workbook.FormulaWithColumn(nameFormula!, newColumn);
            newValuesFormula = workbook.FormulaWithColumn(valuesFormula!, newColumn);
        }

        var newSeries = (OpenXmlCompositeElement)template.CloneNode(true);
        this.SetNextSeriesIndex(newSeries, seriesList);
        this.SetUniqueIds(newSeries);
        this.UpdateSeriesName(newSeries, name, newNameFormula);
        this.UpdateValues(newSeries, valueElementName, values, newValuesFormula);

        workbook?.AppendSeriesColumn(newNameFormula!, newValuesFormula!, name, values);

        chart.InsertAfter(newSeries, seriesList[^1]);
    }

    private static int? PointCount(OpenXmlCompositeElement? numberData)
    {
        var pointCount = numberData?.GetFirstChild<C.PointCount>()?.Val?.Value;
        return pointCount.HasValue ? (int)pointCount.Value : null;
    }

    private bool IsSupportedChart(string localName) => localName is "barChart" or "lineChart" or "pieChart" or "doughnutChart" or "scatterChart" or "areaChart";

    private OpenXmlCompositeElement NumberData(OpenXmlElement series, string localName)
    {
        return series.ChildElements.FirstOrDefault(element => element.LocalName.Equals(localName, StringComparison.Ordinal)) as OpenXmlCompositeElement
            ?? throw new SCException($"The template series does not contain a '{localName}' values element.");
    }

    private int NextSeriesColumn(IReadOnlyList<OpenXmlElement> seriesList, string valuesElementName, Workbook workbook)
    {
        return seriesList
            .Select(series => this.NumberData(series, valuesElementName).GetFirstChild<C.NumberReference>()?.Formula?.Text)
            .Where(formula => !string.IsNullOrWhiteSpace(formula))
            .Select(formula => workbook.FormulaColumnIndex(formula!))
            .DefaultIfEmpty(0)
            .Max() + 1;
    }

    private void ValidateValueCount(OpenXmlElement template, bool isScatterChart, int valueCount)
    {
        int? templatePointCount;
        if (isScatterChart)
        {
            var xValues = this.NumberData(template, "xVal");
            templatePointCount = PointCount(xValues.GetFirstChild<C.NumberReference>()?.NumberingCache)
                ?? PointCount(xValues.GetFirstChild<C.NumberLiteral>());
        }
        else
        {
            var categories = template.GetFirstChild<C.CategoryAxisData>()
                ?? throw new SCException("The template series does not contain chart categories.");
            var categoryCount = categories.StringReference?.StringCache?.PointCount?.Val?.Value
                ?? categories.NumberReference?.NumberingCache?.PointCount?.Val?.Value
                ?? categories.StringLiteral?.PointCount?.Val?.Value
                ?? categories.NumberLiteral?.PointCount?.Val?.Value;
            templatePointCount = categoryCount.HasValue ? (int)categoryCount.Value : null;
        }

        if (templatePointCount == null || templatePointCount.Value != valueCount)
        {
            throw new SCException("The number of values must match the existing category or template X-value count.");
        }
    }

    private void SetNextSeriesIndex(OpenXmlCompositeElement newSeries, IReadOnlyList<OpenXmlElement> seriesList)
    {
        var nextIndex = seriesList
            .Select(series => series.GetFirstChild<C.Index>()?.Val?.Value ?? 0U)
            .DefaultIfEmpty(0U)
            .Max() + 1;
        var nextOrder = seriesList
            .Select(series => series.GetFirstChild<C.Order>()?.Val?.Value ?? 0U)
            .DefaultIfEmpty(0U)
            .Max() + 1;
        (newSeries.GetFirstChild<C.Index>() ?? throw new SCException("The template series does not contain an index.")).Val = nextIndex;
        (newSeries.GetFirstChild<C.Order>() ?? throw new SCException("The template series does not contain an order.")).Val = nextOrder;
    }

    private void SetUniqueIds(OpenXmlElement newSeries)
    {
        foreach (var uniqueId in newSeries.Descendants().Where(element => element.LocalName == "uniqueId"))
        {
            var valueAttribute = uniqueId.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val");
            if (valueAttribute.LocalName == "val")
            {
                uniqueId.SetAttribute(new OpenXmlAttribute(
                    valueAttribute.Prefix,
                    valueAttribute.LocalName,
                    valueAttribute.NamespaceUri,
                    $"{{{Guid.NewGuid():D}}}"));
            }
        }
    }

    private void UpdateSeriesName(OpenXmlCompositeElement series, string name, string? nameFormula)
    {
        var seriesText = series.GetFirstChild<C.SeriesText>()
            ?? throw new SCException("The template series does not contain a name element.");
        if (nameFormula != null && seriesText.StringReference != null)
        {
            var stringReference = seriesText.StringReference;
            stringReference.Formula!.Text = nameFormula;
            var stringCache = stringReference.StringCache ??= new C.StringCache();
            stringCache.RemoveAllChildren<C.PointCount>();
            stringCache.RemoveAllChildren<C.StringPoint>();
            stringCache.Append(new C.PointCount { Val = 1U });
            stringCache.Append(new C.StringPoint(new C.NumericValue(name)) { Index = 0U });
            return;
        }

        seriesText.StringReference?.Remove();
        seriesText.NumericValue = new C.NumericValue(name);
    }

    private void UpdateValues(
        OpenXmlCompositeElement series,
        string valueElementName,
        IReadOnlyList<double> values,
        string? valuesFormula)
    {
        var numberData = this.NumberData(series, valueElementName);
        var numberReference = numberData.GetFirstChild<C.NumberReference>();
        if (valuesFormula != null && numberReference != null)
        {
            numberReference.Formula!.Text = valuesFormula;
            this.UpdateNumericPoints(numberReference.NumberingCache ??= new C.NumberingCache(), values);
            return;
        }

        var numberLiteral = numberData.GetFirstChild<C.NumberLiteral>()
            ?? throw new SCException("The template series does not contain cached or literal numeric values.");
        numberReference?.Remove();
        this.UpdateNumericPoints(numberLiteral, values);
    }

    private void UpdateNumericPoints(OpenXmlCompositeElement numberData, IReadOnlyList<double> values)
    {
        var pointCount = numberData.GetFirstChild<C.PointCount>();
        if (pointCount == null)
        {
            pointCount = new C.PointCount();
            var firstPoint = numberData.GetFirstChild<C.NumericPoint>();
            numberData.InsertBefore(pointCount, firstPoint);
        }

        pointCount.Val = (uint)values.Count;
        numberData.RemoveAllChildren<C.NumericPoint>();
        for (var index = 0; index < values.Count; index++)
        {
            numberData.AppendChild(new C.NumericPoint(new C.NumericValue(values[index].ToString(CultureInfo.InvariantCulture)))
            {
                Index = (uint)index
            });
        }
    }
}