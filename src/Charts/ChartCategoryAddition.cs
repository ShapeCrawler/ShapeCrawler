using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

/// <summary>
///     Appends a category and its series values to a chart.
/// </summary>
internal sealed class ChartCategoryAddition(ChartPart chartPart)
{
    /// <summary>
    ///     Appends a category, keeping chart caches and the embedded workbook in sync.
    /// </summary>
    /// <param name="name">The category name.</param>
    /// <param name="values">One value for each series.</param>
    internal void Add(string name, IReadOnlyList<double> values)
    {
        var plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var charts = plotArea.Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToList();
        if (charts.Count != 1)
        {
            throw new SCException("Adding a category is supported only for non-combination charts.");
        }

        var series = charts[0].ChildElements.Where(element => element.LocalName == "ser").ToList();
        if (series.Count == 0 || series.Count != values.Count)
        {
            throw new SCException("Supply one category value for each existing chart series.");
        }

        var workbook = chartPart.EmbeddedPackagePart is { } package ? new Workbook(package) : null;
        var cells = new List<(string Formula, string Value, bool IsText)>();
        var replacements = new List<(OpenXmlElement Original, OpenXmlElement Updated)>();
        uint? categoryCount = null;
        for (var index = 0; index < series.Count; index++)
        {
            var categories = series[index].GetFirstChild<C.CategoryAxisData>()
                ?? throw new SCException("The chart series does not contain categories.");
            var seriesValues = series[index].GetFirstChild<C.Values>()
                ?? throw new SCException("The chart series does not contain category values.");
            var updatedCategories = (C.CategoryAxisData)categories.CloneNode(true);
            var updatedValues = (C.Values)seriesValues.CloneNode(true);
            var stringData = StringData(updatedCategories);
            var numberData = NumberData(updatedValues);
            var count = Count(stringData);
            if (Count(numberData) != count || (categoryCount.HasValue && categoryCount.Value != count))
            {
                throw new SCException("All series must have matching category and value counts.");
            }

            categoryCount = count;
            var value = values[index].ToString(CultureInfo.InvariantCulture);
            UpdateFormula(updatedCategories.StringReference, workbook, name, true, count, cells);
            UpdateFormula(updatedValues.NumberReference, workbook, value, false, count, cells);
            AppendPoint(stringData, new C.StringPoint(new C.NumericValue(name)) { Index = count }, count);
            AppendPoint(numberData, new C.NumericPoint(new C.NumericValue(value)) { Index = count }, count);
            replacements.Add((categories, updatedCategories));
            replacements.Add((seriesValues, updatedValues));
        }

        workbook?.AppendCategoryCells(cells);
        foreach (var (original, updated) in replacements)
        {
            original.Parent!.ReplaceChild(updated, original);
        }
    }

    private static OpenXmlCompositeElement StringData(C.CategoryAxisData categories)
    {
        return (OpenXmlCompositeElement?)categories.StringReference?.StringCache ?? categories.StringLiteral
            ?? throw new SCException("Adding a category requires single-level text categories with cached or literal data.");
    }

    private static OpenXmlCompositeElement NumberData(C.Values values)
    {
        return (OpenXmlCompositeElement?)values.NumberReference?.NumberingCache ?? values.NumberLiteral
            ?? throw new SCException("Adding a category requires cached or literal series values.");
    }

    private static uint Count(OpenXmlCompositeElement data)
    {
        return data.GetFirstChild<C.PointCount>()?.Val?.Value
            ?? throw new SCException("The chart data does not contain a point count.");
    }

    private static void AppendPoint(OpenXmlCompositeElement data, OpenXmlElement point, uint count)
    {
        data.GetFirstChild<C.PointCount>()!.Val = count + 1;
        var lastPoint = data.ChildElements.LastOrDefault(element => element is C.StringPoint or C.NumericPoint);
        data.InsertAfter(point, lastPoint ?? data.GetFirstChild<C.PointCount>());
    }

    private static void UpdateFormula(
        OpenXmlCompositeElement? reference,
        Workbook? workbook,
        string value,
        bool isText,
        uint count,
        List<(string Formula, string Value, bool IsText)> cells)
    {
        if (reference == null)
        {
            return;
        }

        var formula = reference.GetFirstChild<C.Formula>();
        if (formula == null || workbook == null)
        {
            throw new SCException("Formula-backed chart data requires a worksheet formula and an embedded Excel workbook.");
        }

        var sourceFormula = formula.Text;
        formula.Text = workbook.ExtendedCategoryFormula(sourceFormula, count);
        cells.Add((sourceFormula, value, isText));
    }
}