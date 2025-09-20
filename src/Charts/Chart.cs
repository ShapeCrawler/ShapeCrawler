using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class Chart : IChart
{
    private readonly SeriesCollection seriesCollection;
    private readonly SlideShapeOutline outline;
    private readonly ShapeFill fill;
    private readonly ChartPart chartPart;
    private readonly Categories? categories;
    private readonly XAxis? xAxis;
    private string? chartTitle;
    
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
    }

    public ChartType Type
    {
        get
        {
            var plotArea = this.chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
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

    public string? Title
    {
        get => this.GetTitleOrNull();
        set => this.SetTitle(value);
    }

    public IReadOnlyList<ICategory>? Categories => this.categories;

    public IXAxis? XAxis => this.xAxis;

    public ISeriesCollection SeriesCollection => this.seriesCollection;

    public Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("It is not possible to set the geometry type for the chart shape.");
    }

    public byte[] GetWorksheetByteArray() => new Workbook(this.chartPart.EmbeddedPackagePart!).AsByteArray();

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

    private void SetTitle(string? value)
    {
        this.chartTitle = value;
        var cChart = this.chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var cTitle = cChart.Title;

        if (string.IsNullOrEmpty(value))
        {
            cTitle?.Remove();
            var plotArea = cChart.PlotArea!;
            var autoTitleDeleted = plotArea.GetFirstChild<C.AutoTitleDeleted>();
            if (autoTitleDeleted == null)
            {
                plotArea.InsertAt(new C.AutoTitleDeleted { Val = true }, 0);
            }
            else
            {
                autoTitleDeleted.Val = true;
            }

            return;
        }

        var autoTitleDeletedCheck = cChart.PlotArea!.GetFirstChild<C.AutoTitleDeleted>();
        if (autoTitleDeletedCheck != null)
        {
            autoTitleDeletedCheck.Val = false;
        }

        if (cTitle == null)
        {
            cTitle = new C.Title();
            cChart.InsertAt(cTitle, 0);
        }

        var cChartText = cTitle.GetFirstChild<C.ChartText>() ?? cTitle.AppendChild(new C.ChartText());

        var cRichText = cChartText.GetFirstChild<C.RichText>();
        if (cRichText == null)
        {
            cRichText = cChartText.AppendChild(new C.RichText());
            cRichText.Append(new A.BodyProperties());
            cRichText.Append(new A.ListStyle());
        }

        cRichText.RemoveAllChildren<A.Paragraph>();
        var aParagraph = cRichText.AppendChild(new A.Paragraph());
        aParagraph.Append(new A.Run(new A.Text(value!)));

        if (cTitle.Layout == null)
        {
            cTitle.Append(new C.Layout());
        }

        var cOverlay = cTitle.GetFirstChild<C.Overlay>() ?? cTitle.AppendChild(new C.Overlay());

        cOverlay.Val = false;
    }
}