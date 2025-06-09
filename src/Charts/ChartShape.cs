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

namespace ShapeCrawler.Charts;

internal class Chart(
    SeriesCollection seriesCollection,
    SlideShapeOutline outline,
    ShapeFill fill,
    ChartPart chartPart) : IChart
{
    private string? chartTitle;

    public ChartType Type
    {
        get
        {
            var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
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

    public IShapeOutline Outline => outline;

    public IShapeFill Fill => fill;

    public string? Title
    {
        get
        {
            this.chartTitle = this.GetTitleOrNull();
            return this.chartTitle;
        }
    }

    public IReadOnlyList<ICategory>? Categories => null;

    public virtual IXAxis? XAxis => null;

    public ISeriesCollection SeriesCollection => seriesCollection;

    public Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("It is not possible to set the geometry type for the chart shape.");
    }

    public byte[] GetWorksheetByteArray() => new Workbook(chartPart.EmbeddedPackagePart!).AsByteArray();

    private string? GetTitleOrNull()
    {
        var cTitle = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.Title;
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
            return seriesCollection.First().Name;
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