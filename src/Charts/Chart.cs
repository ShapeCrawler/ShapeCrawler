using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class Chart(
    Shape shape,
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

    public decimal Width
    {
        get => shape.Width;
        set => shape.Width = value;
    }

    public decimal Height
    {
        get => shape.Height;
        set => shape.Height = value;
    }

    public int Id => shape.Id;

    public string Name
    {
        get => shape.Name;
        set => shape.Name = value;
    }

    public string AltText
    {
        get => shape.AltText;
        set => shape.AltText = value;
    }

    public bool Hidden => shape.Hidden;

    public PlaceholderType? PlaceholderType => shape.PlaceholderType;

    public string? CustomData
    {
        get => shape.CustomData;
        set => shape.CustomData = value;
    }

    public ShapeContent ShapeContent => ShapeContent.Chart;

    public IShapeOutline Outline => outline;

    public IShapeFill Fill => fill;

    public ITextBox? TextBox => shape.TextBox;

    public double Rotation => shape.Rotation;

    public string SDKXPath => shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;

    public IPresentation Presentation => shape.Presentation;

    public string? Title
    {
        get
        {
            this.chartTitle = this.GetTitleOrNull();
            return this.chartTitle;
        }
    }

    public IReadOnlyList<ICategory>? Categories => null;

    public IXAxis? XAxis => null;

    public ISeriesCollection SeriesCollection => seriesCollection;

    public Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("Unable to set geometry type for chart.");
    }

    public decimal CornerSize
    {
        get => shape.CornerSize;
        set => shape.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => shape.Adjustments;
        set => shape.Adjustments = value;
    }

    public bool Removable => true;
    
    public decimal X
    {
        get => shape.X;
        set => shape.X = value;
    }

    public decimal Y
    {
        get => shape.Y;
        set => shape.Y = value;
    }

    public byte[] GetWorksheetByteArray() => new Workbook(chartPart.EmbeddedPackagePart!).AsByteArray();

    public void Remove() => shape.Remove();

    public ITable AsTable() => shape.AsTable();

    public IMediaShape AsMedia() => shape.AsMedia();

    public void Duplicate() => shape.Duplicate();

    public void SetText(string text) => shape.SetText(text);

    public void SetImage(string imagePath) => shape.SetImage(imagePath);

    public void SetFontName(string fontName) => shape.SetFontName(fontName);

    public void SetFontSize(decimal fontSize) => shape.SetFontSize(fontSize);

    public void SetFontColor(string colorHex) => shape.SetFontColor(colorHex);

    public void SetVideo(Stream video)
    {
        throw new NotImplementedException();
    }

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