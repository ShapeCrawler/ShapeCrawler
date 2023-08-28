using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideShape;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

internal sealed record SlideChart : IRemoveable, IChart
{
    private readonly ResetableLazy<ICategoryCollection?> categories;
    private readonly Lazy<SCChartType> chartType;
    private readonly Lazy<OpenXmlElement?> firstSeries;
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly Lazy<SCSeriesCollection> series;
    private readonly Lazy<List<double>?> xValues;
    private readonly C.PlotArea cPlotArea;

    // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
    // then collection contains only single item.
    private readonly IEnumerable<OpenXmlElement> cXCharts;

    private string? chartTitle;
    private readonly Shape shape;

    internal SlideChart(SlidePart sdkSlidePart, P.GraphicFrame pGraphicFrame)
    {
        this.pGraphicFrame = pGraphicFrame;
        this.firstSeries = new Lazy<OpenXmlElement?>(this.GetFirstSeries);
        this.xValues = new Lazy<List<double>?>(this.GetXValues);
        this.series = new Lazy<SCSeriesCollection>(this.GetSeries);
        this.categories = new ResetableLazy<ICategoryCollection?>(this.GetCategories);
        this.chartType = new Lazy<SCChartType>(this.GetChartType);

        var cChartReference = this.pGraphicFrame.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !
            .GetFirstChild<C.ChartReference>() !;
        this.ChartPart = (ChartPart)sdkSlidePart.GetPartById(cChartReference.Id!);

        this.cPlotArea = this.ChartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        this.cXCharts = this.cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        this.workbook = this.ChartPart.EmbeddedPackagePart != null ? new ChartSpreadsheet(this.ChartPart.EmbeddedPackagePart) : null;
        this.shape = new Shape(pGraphicFrame);
    }

    public SCChartType Type => this.chartType.Value;

    public SCShapeType ShapeType => SCShapeType.Chart;
    public IAutoShape? AsAutoShape()
    {
        throw new NotImplementedException();
    }

    public string? Title
    {
        get
        {
            this.chartTitle = this.GetTitleOrDefault();
            return this.chartTitle;
        }
    }

    public bool HasTitle
    {
        get
        {
            this.chartTitle ??= this.GetTitleOrDefault();
            return this.chartTitle != null;
        }
    }

    public bool HasCategories => this.categories.Value != null;

    public ISeriesCollection SeriesCollection => this.series.Value;

    public ICategoryCollection? Categories => this.categories.Value;

    public bool HasXValues => this.xValues.Value != null;

    public List<double> XValues
    {
        get
        {
            if (this.xValues.Value == null)
            {
                throw new NotSupportedException(ExceptionMessages.NotXValues);
            }

            return this.xValues.Value;
        }
    }

    public int X
    {
        get => this.shape.X(); 
        set => this.shape.UpdateX(value);
    }

    public int Y
    {
        get => this.shape.Y(); 
        set => this.shape.UpdateY(value);
    }

    public int Width
    {
        get => this.shape.Width(); 
        set => this.shape.UpdateWidth(value);
    }

    public int Height
    {
        get => this.shape.Height(); 
        set => this.shape.UpdateHeight(value);
    }
    
    public int Id => this.shape.Id();
    
    public string Name => this.shape.Name();
    
    public bool Hidden => this.shape.Hidden();
    public bool IsPlaceholder()
    {
        throw new NotImplementedException();
    }

    public IPlaceholder? Placeholder { get; }
    public SCGeometry GeometryType => SCGeometry.Rectangle;
    public string? CustomData { get; set; }

    public byte[] WorkbookByteArray => this.workbook!.BinaryData;

    public SpreadsheetDocument SDKSpreadsheetDocument => this.workbook!.SpreadsheetDocument.Value;
    
    public IAxesManager Axes => this.GetAxes();

    internal ChartSpreadsheet? workbook { get; set; }

    internal ChartPart ChartPart { get; private set; }

    internal void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new NotImplementedException();
    }

    internal string ToJson()
    {
        throw new NotImplementedException();
    }

    private SCChartType GetChartType()
    {
        if (this.cXCharts.Count() > 1)
        {
            return SCChartType.Combination;
        }

        var chartName = this.cXCharts.Single().LocalName;
        Enum.TryParse(chartName, true, out SCChartType enumChartType);

        return enumChartType;
    }
    
    private IAxesManager GetAxes()
    {
        return new SCAxesManager(this.cPlotArea);
    }

    private ICategoryCollection? GetCategories()
    {
        return CategoryCollection.Create(this, this.firstSeries.Value, this.Type);
    }

    private SCSeriesCollection GetSeries()
    {
        return SCSeriesCollection.Create(this, this.cXCharts);
    }

    private string? GetTitleOrDefault()
    {
        var cTitle = this.ChartPart.ChartSpace.GetFirstChild<C.Chart>() !.Title;
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
        if (this.Type == SCChartType.PieChart)
        {
            return ((SCSeriesCollection)this.SeriesCollection).First().Name;
        }

        return null;
    }

    private bool TryGetStaticTitle(C.ChartText chartText, out string? staticTitle)
    {
        staticTitle = null;
        if (this.Type == SCChartType.Combination)
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

    private List<double>? GetXValues()
    {
        var sdkXValues = this.firstSeries.Value?.GetFirstChild<C.XValues>();
        if (sdkXValues?.NumberReference == null)
        {
            return null;
        }

        IEnumerable<double> points =
            ChartReferencesParser.GetNumbersFromCacheOrWorkbook(sdkXValues.NumberReference, this);

        return points.ToList();
    }

    private OpenXmlElement? GetFirstSeries()
    {
        return this.cXCharts.First().ChildElements
            .FirstOrDefault(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
    }

    public void Remove()
    {
        this.pGraphicFrame.Remove();
    }
}
