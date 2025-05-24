using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

// ReSharper disable once InconsistentNaming
internal sealed class OLEObject(Shape shape, SlideShapeOutline outline, ShapeFill fill) : IShape
{
    public decimal Width
    {
        get => shape.Width;
        set=> shape.Width = value;
    }

    public decimal Height
    {
        get=> shape.Height;
        set=> shape.Height = value;
    }
    public int Id => shape.Id;

    public string Name
    {
        get => shape.Name;
        set=> shape.Name = value;
    }

    public string AltText
    {
        get=> shape.AltText;
        set=> shape.AltText = value;
    }
    public bool Hidden => shape.Hidden;
    public PlaceholderType? PlaceholderType => shape.PlaceholderType;

    public string? CustomData
    {
        get=> shape.CustomData;
        set=> shape.CustomData = value;
    }
    public ShapeContent ShapeContent => ShapeContent.OLEObject;

    public IShapeOutline Outline => outline;

    public IShapeFill Fill => fill;
    public ITextBox? TextBox => shape.TextBox;
    public double Rotation => shape.Rotation;
    public string SDKXPath => shape.SDKXPath;
    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;
    public IPresentation Presentation => shape.Presentation;

    public bool Removable => true;

    public void Remove() => shape.Remove();
    public ITable AsTable()=> shape.AsTable();

    public IMediaShape AsMedia()=> shape.AsMedia();

    public void Duplicate()=> shape.Duplicate();

    public void SetText(string text)=> shape.SetText(text);

    public void SetImage(string imagePath)=> shape.SetImage(imagePath);

    public decimal X
    {
        get=> shape.X;
        set=> shape.X = value;
    }

    public decimal Y
    {
        get=> shape.Y;
        set=> shape.Y = value;
    }

    public Geometry GeometryType
    {
        get=> shape.GeometryType;
        set=> shape.GeometryType = value;
    }

    public decimal CornerSize
    {
        get=> shape.CornerSize;
        set=> shape.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get=> shape.Adjustments;
        set=> shape.Adjustments = value;
    }
}