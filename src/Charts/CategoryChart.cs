using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

internal class CategoryChart : IChart
{
    private readonly IChart chart;

    internal CategoryChart(IChart chart, ChartPart chartPart)
    {
        this.chart = chart;
        this.Categories = new Categories(chartPart);
    }

    public IReadOnlyList<ICategory>? Categories { get; }

    #region Composition Properties

    public decimal X
    {
        get => this.chart.X;
        set => this.chart.X = value;
    }

    public decimal Y
    {
        get => this.chart.Y;
        set => this.chart.Y = value;
    }

    public Geometry GeometryType
    {
        get => this.chart.GeometryType;
        set => this.chart.GeometryType = value;
    }

    public decimal CornerSize
    {
        get => this.chart.CornerSize;
        set => this.chart.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => this.chart.Adjustments;
        set => this.chart.Adjustments = value;
    }

    public decimal Width
    {
        get => this.chart.Width;
        set => this.chart.Width = value;
    }

    public decimal Height
    {
        get => this.chart.Height;
        set => this.chart.Height = value;
    }

    public int Id => this.chart.Id;

    public string Name
    {
        get => this.chart.Name;
        set => this.chart.Name = value;
    }

    public string AltText
    {
        get => this.chart.AltText;
        set => this.chart.AltText = value;
    }

    public bool Hidden => this.chart.Hidden;

    public PlaceholderType? PlaceholderType => this.chart.PlaceholderType;

    public string? CustomData
    {
        get => this.chart.CustomData;
        set => this.chart.CustomData = value;
    }

    public ShapeContent ShapeContent => this.chart.ShapeContent;

    public IShapeOutline? Outline => this.chart.Outline;

    public IShapeFill? Fill => this.chart.Fill;

    public ITextBox? TextBox => this.chart.TextBox;
   
    public double Rotation => this.chart.Rotation;
    
    public string SDKXPath => this.chart.SDKXPath;
    
    public OpenXmlElement SDKOpenXmlElement => this.chart.SDKOpenXmlElement;

    public IPresentation Presentation => this.chart.Presentation;

    public ChartType Type => this.chart.Type;

    public string? Title => this.chart.Title;

    public IXAxis? XAxis => this.chart.XAxis;

    public ISeriesCollection SeriesCollection => this.chart.SeriesCollection;

    #endregion Composition Properties

    #region Composition Methods

    public void Remove() => this.chart.Remove();

    public ITable AsTable() => this.chart.AsTable();

    public IMediaShape AsMedia() => this.chart.AsMedia();

    public void Duplicate() => this.chart.Duplicate();

    public byte[] GetWorksheetByteArray() => this.chart.GetWorksheetByteArray();

    #endregion Composition Methods
}