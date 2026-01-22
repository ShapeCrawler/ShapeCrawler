using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class Chart : IChart, IBarChart, IColumnChart, ILineChart, IPieChart, IScatterChart, IBubbleChart,
    IAreaChart
{
    private readonly SeriesCollection seriesCollection;
    private readonly SlideShapeOutline outline;
    private readonly ShapeFill fill;
    private readonly ChartPart chartPart;
    private readonly Categories? categories;
    private readonly XAxis? xAxis;
    private readonly Lazy<ChartTitle> chartTitle;

    internal Chart(
        SeriesCollection seriesCollection,
        SlideShapeOutline outline,
        ShapeFill fill,
        ChartPart chartPart,
        Categories categories,
        XAxis xAxis)
    {
        this.seriesCollection = seriesCollection;
        this.outline = outline;
        this.fill = fill;
        this.chartPart = chartPart;
        this.categories = categories;
        this.xAxis = xAxis;
        this.chartTitle = new Lazy<ChartTitle>(() => new ChartTitle(chartPart, this.Type, this.SeriesCollection, new ChartTitleAlignment(chartPart)));
    }

    internal Chart(
        SeriesCollection seriesCollection,
        SlideShapeOutline outline,
        ShapeFill fill,
        ChartPart chartPart,
        XAxis xAxis)
    {
        this.seriesCollection = seriesCollection;
        this.outline = outline;
        this.fill = fill;
        this.chartPart = chartPart;
        this.xAxis = xAxis;
        this.chartTitle = new Lazy<ChartTitle>(() => new ChartTitle(chartPart, this.Type, this.SeriesCollection, new ChartTitleAlignment(chartPart)));
    }

    internal Chart(
        SeriesCollection seriesCollection,
        SlideShapeOutline outline,
        ShapeFill fill,
        ChartPart chartPart,
        Categories categories)
    {
        this.seriesCollection = seriesCollection;
        this.outline = outline;
        this.fill = fill;
        this.chartPart = chartPart;
        this.categories = categories;
        this.chartTitle = new Lazy<ChartTitle>(() => new ChartTitle(chartPart, this.Type, this.SeriesCollection, new ChartTitleAlignment(chartPart)));
    }

    public ChartType Type
    {
        get
        {
            var plotArea = this.chartPart.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
            var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            if (cXCharts.Count() > 1)
            {
                return ChartType.Combination;
            }

            var chartName = cXCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType enumChartType);

            return enumChartType;
        }
    }

    public IShapeOutline Outline => this.outline;

    public IShapeFill Fill => this.fill;

    public IChartTitle? Title => this.chartTitle.Value;

    public IReadOnlyList<ICategory>? Categories => this.categories;

    public IXAxis? XAxis => this.xAxis;

    public ISeriesCollection SeriesCollection => this.seriesCollection;

    public Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("It is not possible to set the geometry type for the chart shape.");
    }

    public byte[] GetWorksheetByteArray() => new Workbook(this.chartPart.EmbeddedPackagePart!).AsByteArray();
}