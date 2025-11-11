using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
///     Represents a chart X-axis.
/// </summary>
public interface IXAxis
{
    /// <summary>
    ///     Gets axis values.
    /// </summary>
    double[] Values { get; }

    /// <summary>
    ///     Gets or sets axis minimum value.
    /// </summary>
    double Minimum { get; set; }

    /// <summary>
    ///     Gets or sets axis maximum value.
    /// </summary>
    double Maximum { get; set; }

    /// <summary>
    ///     Gets or sets the X-axis title text. Returns <c>null</c> if not set.
    /// </summary>
    string? Title { get; set; }
}

internal class XAxis(ChartPart chartPart) : IXAxis
{
    public double[] Values
    {
        get
        {
            var cXValues = this.FirstSeries().GetFirstChild<C.XValues>()!;

            if (cXValues.NumberReference!.NumberingCache != null)
            {
                var cNumericValues = cXValues.NumberReference.NumberingCache.Descendants<C.NumericValue>();
                var cachedPointValues = new List<double>(cNumericValues.Count());
                foreach (var numericValue in cNumericValues)
                {
                    var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    var roundNumber = Math.Round(number, 1);
                    cachedPointValues.Add(roundNumber);
                }

                return [.. cachedPointValues];
            }

            return
            [
                .. new Workbook(chartPart.EmbeddedPackagePart!).FormulaValues(cXValues.NumberReference.Formula!.Text)
            ];
        }
    }

    public double Minimum
    {
        get
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!
                .Scaling!;
            var cMin = cScaling.MinAxisValue;

            return cMin == null ? 0 : cMin.Val!;
        }

        set
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!
                .Scaling!;
            cScaling.MinAxisValue = new C.MinAxisValue { Val = value };
        }
    }

    public double Maximum
    {
        get
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!
                .Scaling!;
            var cMax = cScaling.MaxAxisValue;
            const double defaultMax = 6;

            return cMax == null ? defaultMax : cMax.Val!;
        }

        set
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!
                .Scaling!;
            cScaling.MaxAxisValue = new C.MaxAxisValue { Val = value };
        }
    }

    public string? Title
    {
        get
        {
            var axis = this.GetXAxisElement();
            if (axis == null)
            {
                return null;
            }

            var cTitle = axis.GetFirstChild<C.Title>();
            return cTitle?.Descendants<A.Text>().FirstOrDefault()?.Text;
        }

        set
        {
            var axis = this.GetXAxisElement();
            if (axis == null)
            {
                return;
            }

            if (string.IsNullOrEmpty(value))
            {
                axis.GetFirstChild<C.Title>()?.Remove();
                return;
            }

            var cTitle = GetOrCreateTitle(axis);
            var chartText = GetOrCreateChartText(cTitle);
            var richText = GetOrCreateRichText(chartText);

            UpdateRichTextContent(richText, value!);
            EnsureOverlayDisabled(cTitle);
        }
    }

    private static C.Title GetOrCreateTitle(OpenXmlCompositeElement axis)
    {
        var title = axis.GetFirstChild<C.Title>();
        if (title != null)
        {
            return title;
        }

        title = new C.Title();
        var insertBefore = axis.Elements<OpenXmlElement>()
            .FirstOrDefault(IsTitleInsertionBoundary);

        if (insertBefore != null)
        {
            axis.InsertBefore(title, insertBefore);
        }
        else
        {
            axis.AppendChild(title);
        }

        return title;
    }

    private static C.ChartText GetOrCreateChartText(C.Title title) =>
        title.GetFirstChild<C.ChartText>() ?? title.AppendChild(new C.ChartText());

    private static C.RichText GetOrCreateRichText(C.ChartText chartText)
    {
        var richText = chartText.GetFirstChild<C.RichText>();
        if (richText != null)
        {
            return richText;
        }

        richText = chartText.AppendChild(new C.RichText());
        richText.Append(new A.BodyProperties());
        richText.Append(new A.ListStyle());

        return richText;
    }

    private static void UpdateRichTextContent(C.RichText richText, string text)
    {
        richText.RemoveAllChildren<A.Paragraph>();
        var paragraph = richText.AppendChild(new A.Paragraph());
        paragraph.AppendChild(new A.Run(new A.Text(text)));
    }

    private static void EnsureOverlayDisabled(C.Title title)
    {
        var overlay = title.GetFirstChild<C.Overlay>() ?? title.AppendChild(new C.Overlay());
        overlay.Val = false;
    }

    private static bool IsTitleInsertionBoundary(OpenXmlElement element) =>
        element switch
        {
            C.NumberingFormat => true,
            C.MajorTickMark => true,
            C.MinorTickMark => true,
            C.TickLabelPosition => true,
            C.CrossingAxis => true,
            C.Crosses => true,
            C.CrossBetween => true,
            C.CrossesAt => true,
            C.Layout => true,
            C.ShapeProperties => true,
            C.TextProperties => true,
            _ => false,
        };

    private OpenXmlElement FirstSeries()
    {
        var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        return cXCharts.First().ChildElements
            .First(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
    }

    private OpenXmlCompositeElement? GetXAxisElement()
    {
        var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        var categoryAxis = plotArea.Elements<C.CategoryAxis>()
                               .FirstOrDefault(a => a.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom)
                           ?? plotArea.Elements<C.CategoryAxis>().FirstOrDefault();
        if (categoryAxis != null)
        {
            return categoryAxis;
        }

        return plotArea.Elements<C.ValueAxis>()
                   .FirstOrDefault(a => a.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom)
               ?? plotArea.Elements<C.ValueAxis>().FirstOrDefault();
    }
}