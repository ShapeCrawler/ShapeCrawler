using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

/// <summary>
///     Adds a series to a template bar or column chart.
/// </summary>
internal sealed class ChartSeriesAddition(ChartPart chartPart, IEnumerable<OpenXmlElement> chartElements)
{
    /// <summary>
    ///     Adds a series using the chart's existing categories and template formatting.
    /// </summary>
    /// <param name="name">The name of the series.</param>
    /// <param name="values">One value for each existing category.</param>
    internal void Add(string name, IReadOnlyList<double> values)
    {
        var chartList = chartElements.ToList();
        if (chartList.Count != 1 || chartList[0] is not C.BarChart barChart)
        {
            throw new SCException("Adding a series is supported only for non-combination 2-D bar and column charts.");
        }

        var seriesList = barChart.Elements<C.BarChartSeries>().ToList();
        if (seriesList.Count == 0)
        {
            throw new SCException("Cannot add a series to a bar or column chart without an existing template series.");
        }

        var template = seriesList[0];
        var nameFormula = template.SeriesText?.StringReference?.Formula?.Text;
        var valuesFormula = template.GetFirstChild<C.Values>()?.NumberReference?.Formula?.Text;
        if (string.IsNullOrWhiteSpace(nameFormula)
            || string.IsNullOrWhiteSpace(valuesFormula)
            || chartPart.EmbeddedPackagePart == null)
        {
            throw new SCException("Adding a series requires an existing template series linked to an embedded Excel workbook.");
        }

        var categoryData = template.GetFirstChild<C.CategoryAxisData>();
        var categoryCount = categoryData?.StringReference?.StringCache?.PointCount?.Val
            ?? categoryData?.NumberReference?.NumberingCache?.PointCount?.Val;
        if (categoryCount == null || categoryCount.Value != values.Count)
        {
            throw new SCException("The number of values must match the existing chart category count.");
        }

        var workbook = new Workbook(chartPart.EmbeddedPackagePart);
        var newColumn = seriesList
            .Select(series => series.GetFirstChild<C.Values>()?.NumberReference?.Formula?.Text)
            .Where(formula => !string.IsNullOrWhiteSpace(formula))
            .Select(formula => workbook.FormulaColumnIndex(formula!))
            .DefaultIfEmpty(0)
            .Max() + 1;
        var newNameFormula = workbook.FormulaWithColumn(nameFormula!, newColumn);
        var newValuesFormula = workbook.FormulaWithColumn(valuesFormula!, newColumn);

        var newSeries = (C.BarChartSeries)template.CloneNode(true);
        newSeries.Index!.Val = seriesList.Max(series => series.Index?.Val?.Value ?? 0U) + 1;
        newSeries.Order!.Val = seriesList.Max(series => series.Order?.Val?.Value ?? 0U) + 1;
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

        var seriesText = newSeries.SeriesText!;
        seriesText.StringReference!.Formula!.Text = newNameFormula;
        var stringCache = seriesText.StringReference.StringCache ??= new C.StringCache();
        stringCache.RemoveAllChildren<C.PointCount>();
        stringCache.RemoveAllChildren<C.StringPoint>();
        stringCache.Append(new C.PointCount { Val = 1U });
        stringCache.Append(new C.StringPoint(new C.NumericValue(name)) { Index = 0U });

        var numberReference = newSeries.GetFirstChild<C.Values>()!.NumberReference!;
        numberReference.Formula!.Text = newValuesFormula;
        var numberCache = numberReference.NumberingCache ??= new C.NumberingCache();
        numberCache.RemoveAllChildren<C.PointCount>();
        numberCache.RemoveAllChildren<C.NumericPoint>();
        numberCache.Append(new C.PointCount { Val = (uint)values.Count });
        for (var index = 0; index < values.Count; index++)
        {
            numberCache.Append(new C.NumericPoint(new C.NumericValue(values[index].ToString(CultureInfo.InvariantCulture)))
            {
                Index = (uint)index
            });
        }

        // Update the embedded spreadsheet before appending chart XML so a failed workbook update leaves the chart unchanged.
        workbook.AppendSeriesColumn(newNameFormula, newValuesFormula, name, values);
        barChart.InsertAfter(newSeries, seriesList[^1]);
    }
}