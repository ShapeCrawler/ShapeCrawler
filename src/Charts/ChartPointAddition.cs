using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

/// <summary>
///     Appends a numeric chart point while synchronizing its worksheet range.
/// </summary>
/// <param name="chartPart">The chart part owning the data.</param>
/// <param name="values">The numeric data element.</param>
internal sealed class ChartPointAddition(ChartPart chartPart, OpenXmlCompositeElement? values)
{
    /// <summary>
    ///     Appends a numeric value to the chart cache or literal data and the embedded workbook.
    /// </summary>
    /// <param name="value">The point value.</param>
    internal void Add(double value)
    {
        var reference = values?.GetFirstChild<C.NumberReference>();
        var data = (OpenXmlCompositeElement?)reference?.NumberingCache ?? values?.GetFirstChild<C.NumberLiteral>()
            ?? throw new SCException("Adding a point requires cached or literal numeric chart data.");
        var nextIndex = data.Elements<C.NumericPoint>().Select(point => point.Index?.Value ?? 0U)
            .Select(index => index + 1).DefaultIfEmpty(0U).Max();
        var count = Math.Max(data.GetFirstChild<C.PointCount>()?.Val?.Value ?? 0U, nextIndex);
        var text = value.ToString(CultureInfo.InvariantCulture);
        this.UpdateWorksheet(reference, count, text);

        var pointCount = data.GetFirstChild<C.PointCount>();
        if (pointCount == null)
        {
            pointCount = new C.PointCount();
            data.AddChild(pointCount);
        }

        pointCount.Val = count + 1;
        var lastPoint = data.Elements<C.NumericPoint>().LastOrDefault();
        data.InsertAfter(new C.NumericPoint(new C.NumericValue(text)) { Index = count }, (OpenXmlElement?)lastPoint ?? pointCount);
    }

    private void UpdateWorksheet(C.NumberReference? reference, uint count, string value)
    {
        if (reference == null)
        {
            return;
        }

        var formula = reference.Formula;
        if (formula == null || chartPart.EmbeddedPackagePart == null)
        {
            throw new SCException("Formula-backed chart data requires a worksheet formula and an embedded Excel workbook.");
        }

        var workbook = new Workbook(chartPart.EmbeddedPackagePart);
        var updatedFormula = workbook.ExtendedPointFormula(formula.Text, count);
        workbook.AppendPointCells([(formula.Text, value, false)]);
        formula.Text = updatedFormula;
    }
}