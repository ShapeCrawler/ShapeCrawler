using System;
using System.Collections.Generic;
using System.Globalization;
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

internal sealed class ScatterChart : IScatterChart
{
    private readonly Chart chart;

    internal ScatterChart(Chart chart, ChartPart chartPart)
    {
        this.chart = chart;
        XAxis = new XAxis(chartPart);
    }
    
    public ChartType Type => ChartType.ScatterChart;
    
    public IXAxis XAxis { get; }

    #region Composition Properties

    public decimal X
    {
        get => chart.X;
        set => chart.X = value;
    }

    public decimal Y
    {
        get => chart.Y;
        set => chart.Y = value;
    }

    public Geometry GeometryType
    {
        get => chart.GeometryType;
        set => chart.GeometryType = value;
    }

    public decimal Width
    {
        get => chart.Width;
        set => chart.Width = value;
    }

    public decimal Height
    {
        get => chart.Height;
        set => chart.Height = value;
    }

    public int Id => chart.Id;

    public string Name
    {
        get => chart.Name;
        set => chart.Name = value;
    }

    public string AltText
    {
        get => chart.AltText;
        set => chart.AltText = value;
    }

    public bool Hidden => chart.Hidden;
    public PlaceholderType? PlaceholderType => chart.PlaceholderType;

    public string? CustomData
    {
        get => chart.CustomData;
        set => chart.CustomData = value;
    }

    public ShapeContent ShapeContent => chart.ShapeContent;
    public IShapeOutline? Outline => chart.Outline;
    public IShapeFill? Fill => chart.Fill;
    public ITextBox? TextBox => chart.TextBox;
    public double Rotation => chart.Rotation;
    public string SDKXPath => chart.SDKXPath;
    public OpenXmlElement SDKOpenXmlElement => chart.SDKOpenXmlElement;
    public IPresentation Presentation => chart.Presentation;

    public decimal CornerSize
    {
        get => chart.CornerSize;
        set => chart.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => chart.Adjustments;
        set => chart.Adjustments = value;
    }

    public bool HasTitle => chart.HasTitle;
    public string? Title { get; }
    public bool HasCategories { get; }
    public IReadOnlyList<ICategory> Categories => chart.Categories;
    public ISeriesCollection SeriesCollection => chart.SeriesCollection;

    #endregion Composition Properties

    public IScatterChart AsScatterChart() => this;

    #region Composition Methods

    public void Remove() => chart.Remove();

    public ITable AsTable() => chart.AsTable();

    public IMediaShape AsMedia() => chart.AsMedia();

    public void Duplicate() => chart.Duplicate();

    public byte[] GetSpreadsheetByteArray() => chart.GetSpreadsheetByteArray();

    #endregion Composition Methods
}