using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using C = DocumentFormat.OpenXml.Drawing.Charts;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a chart series.
/// </summary>
public interface ISeries
{
    /// <summary>
    ///     Gets or sets series name.
    /// </summary>
    string Name { get; set; }

    /// <summary>
    ///     Gets chart type.
    /// </summary>
    ChartType Type { get; }

    /// <summary>
    ///     Gets the collection of chart points.
    /// </summary>
    IReadOnlyList<IChartPoint> Points { get; }

    /// <summary>
    ///     Gets the collection of X-values points of the series.
    ///     For a scatter chart, assigning <see cref="IChartPoint.Value"/> updates the X coordinate
    ///     in both the chart cache and the embedded worksheet.
    ///     Returns <see langword="null"/> when the series doesn't support X-values.
    /// </summary>
    IReadOnlyList<IChartPoint>? XPoints { get; }

    /// <summary>
    ///     Gets the collection of bubble size points of the series.
    ///     Returns <see langword="null"/> when the series doesn't support bubble size values.
    /// </summary>
    IReadOnlyList<IChartPoint>? BubbleSizePoints { get; }

    /// <summary>
    ///     Gets a value indicating whether chart has name.
    /// </summary>
    bool HasName { get; }
}

internal sealed class Series : ISeries
{
    private readonly ChartPart chartPart;
    private readonly OpenXmlElement cSer;

    internal Series(ChartPart sdkChartPart, OpenXmlElement cSer, ChartType type)
    {
        this.chartPart = sdkChartPart;
        this.cSer = cSer;
        this.Type = type;
        this.Points = new ChartPoints(this.chartPart, this.cSer);
        this.XPoints = type is ChartType.ScatterChart or ChartType.BubbleChart
            ? new SeriesXPoints(this.chartPart, this.cSer)
            : null;
        this.BubbleSizePoints = type is ChartType.BubbleChart
            ? new SeriesBubbleSizePoints(this.chartPart, this.cSer)
            : null;
    }

    public ChartType Type { get; }

    public IReadOnlyList<IChartPoint> Points { get; }

    public IReadOnlyList<IChartPoint>? XPoints { get; }

    public IReadOnlyList<IChartPoint>? BubbleSizePoints { get; }

    public bool HasName
    {
        get
        {
            var cSeriesText = this.cSer.GetFirstChild<C.SeriesText>();
            return cSeriesText?.NumericValue != null || cSeriesText?.StringReference != null;
        }
    }

    public string Name
    {
        get => this.ParseName();
        set => this.SetName(value);
    }

    private static C.StringCache UpdatedStringCache(C.StringCache? currentCache, string value)
    {
        var updatedCache = (C.StringCache?)currentCache?.CloneNode(true) ?? new C.StringCache();
        updatedCache.RemoveAllChildren<C.PointCount>();
        updatedCache.RemoveAllChildren<C.StringPoint>();
        updatedCache.AddChild(new C.PointCount { Val = 1U });
        updatedCache.AddChild(
            new C.StringPoint(new C.NumericValue(value))
            {
                Index = 0U
            });

        return updatedCache;
    }

    private string ParseName()
    {
        var cSeriesText = this.cSer.GetFirstChild<C.SeriesText>();
        var cStringReference = cSeriesText?.StringReference;
        if (cStringReference != null)
        {
            var cachedName = cStringReference.StringCache?
                .Elements<C.StringPoint>()
                .OrderBy(point => point.Index?.Value ?? uint.MaxValue)
                .Select(point => point.NumericValue?.InnerText)
                .FirstOrDefault(name => name != null);
            if (cachedName != null)
            {
                return cachedName;
            }

            var formula = cStringReference.Formula?.Text;
            if (formula != null && this.chartPart.EmbeddedPackagePart != null)
            {
                return new Workbook(this.chartPart.EmbeddedPackagePart).FormulaText(formula);
            }
        }

        return cSeriesText?.NumericValue?.InnerText
            ?? throw new SCException(
                $"Series does not have name. Use {nameof(this.HasName)} property to check if series has name.");
    }

    private void SetName(string value)
    {
        var cSeriesText = this.cSer.GetFirstChild<C.SeriesText>();
        if (cSeriesText == null)
        {
            ((OpenXmlCompositeElement)this.cSer).AddChild(
                new C.SeriesText(new C.NumericValue(value)));
            return;
        }

        var cStringReference = cSeriesText.StringReference;
        if (cStringReference == null)
        {
            cSeriesText.NumericValue ??= new C.NumericValue();
            cSeriesText.NumericValue.Text = value;
            return;
        }

        var updatedCache = UpdatedStringCache(cStringReference.StringCache, value);
        var formula = cStringReference.Formula?.Text;
        if (formula != null && this.chartPart.EmbeddedPackagePart != null)
        {
            new Workbook(this.chartPart.EmbeddedPackagePart).UpdateFormulaCell(formula, value);
        }

        cStringReference.StringCache = updatedCache;
        cSeriesText.NumericValue?.Remove();
    }
}