using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

internal sealed class Chart : Shape, IChart
{
    private readonly SeriesCollection seriesCollection;
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly ChartPart chartPart;

    // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
    // then collection contains only single item.
    private readonly IEnumerable<OpenXmlElement> cXCharts;

    private string? chartTitle;

    internal Chart(ChartPart chartPart, P.GraphicFrame pGraphicFrame)
        : base(pGraphicFrame)
    {
        this.chartPart = chartPart;
        this.pGraphicFrame = pGraphicFrame;
        var plotArea1 = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        this.cXCharts = plotArea1.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        var pShapeProperties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() !;
        this.Outline = new SlideShapeOutline(pShapeProperties);
        this.Fill = new ShapeFill(pShapeProperties);
        this.seriesCollection = new SeriesCollection(
            chartPart,
            plotArea1.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal)));
    }

    public ChartType Type
    {
        get
        {
            if (this.cXCharts.Count() > 1)
            {
                return ChartType.Combination;
            }

            var chartName = this.cXCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType enumChartType);

            return enumChartType;
        }
    }

    public override ShapeContent ShapeContent => ShapeContent.Chart;

    public override IShapeOutline Outline { get; }

    public override IShapeFill Fill { get; }

    public string? Title
    {
        get
        {
            this.chartTitle = this.GetTitleOrNull();
            return this.chartTitle;
        }
    }

    public IReadOnlyList<ICategory>? Categories { get; }

    public IXAxis? XAxis { get; }

    public ISeriesCollection SeriesCollection => this.seriesCollection;

    public override Geometry GeometryType => Geometry.Rectangle;

    public override bool Removable => true;

    public byte[] GetWorksheetByteArray() => new Workbook(this.chartPart.EmbeddedPackagePart!).AsByteArray();

    public override void Remove() => this.pGraphicFrame.Remove();

    private string? GetTitleOrNull()
    {
        var cTitle = this.chartPart.ChartSpace.GetFirstChild<C.Chart>() !.Title;
        if (cTitle == null)
        {
            // chart has not title
            return null;
        }

        var cChartText = cTitle.ChartText;
        bool staticAvailable = this.TryGetStaticTitle(cChartText!, out var staticTitle);
        if (staticAvailable)
        {
            return staticTitle;
        }

        // Dynamic title
        if (cChartText != null)
        {
            return cChartText.Descendants<C.StringPoint>().Single().InnerText;
        }

        // PieChart uses only one series for view.
        // However, it can have store multiple series data in the spreadsheet.
        if (this.Type == ChartType.PieChart)
        {
            return this.seriesCollection.First().Name;
        }

        return null;
    }

    private bool TryGetStaticTitle(C.ChartText chartText, out string? staticTitle)
    {
        staticTitle = null;
        if (this.Type == ChartType.Combination)
        {
            staticTitle = chartText.RichText!.Descendants<A.Text>().Select(t => t.Text)
                .Aggregate((t1, t2) => t1 + t2);
            return true;
        }

        var rRich = chartText?.RichText;
        if (rRich != null)
        {
            staticTitle = rRich.Descendants<A.Text>().Select(t => t.Text).Aggregate((t1, t2) => t1 + t2);
            return true;
        }

        return false;
    }
}