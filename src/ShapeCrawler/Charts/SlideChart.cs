using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

internal sealed class SlideChart : Shape, IChart, IRemoveable
{
    private readonly SCChartType chartType;
    private readonly Lazy<OpenXmlElement?> firstSeries;
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly ChartPart sdkChartPart;
    private readonly C.PlotArea cPlotArea;

    // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
    // then collection contains only single item.
    private readonly IEnumerable<OpenXmlElement> cXCharts;

    private string? chartTitle;

    internal SlideChart(SlidePart sdkSlidePart, P.GraphicFrame pGraphicFrame, ChartPart sdkChartPart)
        : base(pGraphicFrame)
    {
        this.pGraphicFrame = pGraphicFrame;
        this.sdkChartPart = sdkChartPart;
        this.firstSeries = new Lazy<OpenXmlElement?>(this.GetFirstSeries);
        this.cPlotArea = sdkChartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        this.cXCharts = this.cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        
        var pShapeProperties = sdkChartPart.ChartSpace.GetFirstChild<C.ShapeProperties>()!;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pShapeProperties);
        this.Fill = new SlideShapeFill(sdkSlidePart, pShapeProperties, false);
        this.SeriesList = new SeriesList(sdkChartPart,
            this.cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal)));
    }

    public SCChartType Type
    {
        get
        {
            if (this.cXCharts.Count() > 1)
            {
                return SCChartType.Combination;
            }

            var chartName = this.cXCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out SCChartType enumChartType);

            return enumChartType;
        }
    }

    public override SCShapeType ShapeType => SCShapeType.Chart;
    public override IShapeOutline Outline { get; }
    public override IShapeFill Fill { get; }

    public bool HasTitle
    {
        get
        {
            this.chartTitle ??= this.GetTitleOrDefault();
            return this.chartTitle != null;
        }
    }
    public string? Title
    {
        get
        {
            this.chartTitle = this.GetTitleOrDefault();
            return this.chartTitle;
        }
    }
    public bool HasCategories => false;
    public IReadOnlyList<ICategory> Categories => throw new SCException($"Chart does not have categories. Use {nameof(IChart.HasCategories)} property to check if chart categories are available.");
    public ISeriesList SeriesList { get; }
    public bool HasXValues => this.ParseXValues() != null;

    public List<double> XValues
    {
        get
        {
            if (this.ParseXValues() == null)
            {
                throw new NotSupportedException(ExceptionMessages.NotXValues);
            }

            return this.ParseXValues()!;
        }
    }
    public override SCGeometry GeometryType => SCGeometry.Rectangle;
    public byte[] ExcelBookByteArray() => new ExcelBook(this.sdkChartPart).AsByteArray();
    public IAxesManager Axes => this.GetAxes();
    internal ExcelBook? workbook { get; set; }
    private IAxesManager GetAxes()
    {
        return new SCAxesManager(this.cPlotArea);
    }

    private string? GetTitleOrDefault()
    {
        var cTitle = this.sdkChartPart.ChartSpace.GetFirstChild<C.Chart>() !.Title;
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
            return ((SeriesList)this.SeriesList).First().Name;
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

    private List<double>? ParseXValues()
    {
        var cXValues = this.firstSeries.Value?.GetFirstChild<C.XValues>();
        if (cXValues?.NumberReference == null)
        {
            return null;
        }

        if (cXValues.NumberReference.NumberingCache != null)
        {
            // From cache
            var cNumericValues = cXValues.NumberReference.NumberingCache.Descendants<C.NumericValue>();
            var cachedPointValues = new List<double>(cNumericValues.Count());
            foreach (var numericValue in cNumericValues)
            {
                var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                var roundNumber = Math.Round(number, 1);
                cachedPointValues.Add(roundNumber);
            }

            return cachedPointValues;
        }

        // From Spreadsheet
        return new ExcelBook(this.sdkChartPart).FormulaValues(cXValues.NumberReference.Formula!);
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