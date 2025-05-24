using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Charts;

internal sealed class AxisChart(IChart chart, XAxis xAxis) : IChart
{
    public IXAxis XAxis => xAxis;

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
    
    public ChartType Type => chart.Type;

    public string? Title => chart.Title;
    
    public IReadOnlyList<ICategory>? Categories => chart.Categories;

    public ISeriesCollection SeriesCollection => chart.SeriesCollection;

    public void Remove() => chart.Remove();

    public ITable AsTable() => chart.AsTable();

    public IMediaShape AsMedia() => chart.AsMedia();

    public void Duplicate() => chart.Duplicate();

    public void SetText(string text) => chart.SetText(text);

    public void SetImage(string imagePath) => chart.SetImage(imagePath);

    public byte[] GetWorksheetByteArray() => chart.GetWorksheetByteArray();
}