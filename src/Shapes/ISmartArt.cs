using DocumentFormat.OpenXml;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a SmartArt graphic.
/// </summary>
public interface ISmartArt : IShape
{
    /// <summary>
    ///     Gets the collection of nodes in the SmartArt graphic.
    /// </summary>
    ISmartArtNodeCollection Nodes { get; }
}

internal class SmartArt(Shape shape, SmartArtNodeCollection nodeCollection): ISmartArt
{
    // internal SmartArt(OpenXmlElement pShapeTreeElement) 
    //     : base(pShapeTreeElement)
    // {
    //     this.Nodes = new SmartArtNodeCollection();
    // }
    
    public ISmartArtNodeCollection Nodes => nodeCollection;

    public decimal X
    {
        get => shape.X;
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

    public decimal Width
    {
        get=> shape.Width;
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
        get=> shape.Name;
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
    public ShapeContent ShapeContent => shape.ShapeContent;
    public IShapeOutline? Outline => shape.Outline;
    public IShapeFill? Fill => shape.Fill;
    public ITextBox? TextBox => shape.TextBox;
    public double Rotation => shape.Rotation;
    public string SDKXPath => shape.SDKXPath;
    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;
    public IPresentation Presentation => shape.Presentation;
    public void Remove() => shape.Remove();

    public ITable AsTable() => shape.AsTable();
    
    public IMediaShape AsMedia()=> shape.AsMedia();
    
    public void Duplicate()=> shape.Duplicate();

    public void SetText(string text)=> shape.SetText(text);

    public void SetImage(string imagePath)=> shape.SetImage(imagePath);
}
