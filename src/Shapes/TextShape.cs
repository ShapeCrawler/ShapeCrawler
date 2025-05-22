using DocumentFormat.OpenXml;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal class TextShape(Shape shape, TextBox textBox) : IShape
{
    public void SetText(string text) => textBox.SetText(text);

    #region Composition
    
    public virtual decimal X
    {
        get => shape.X;
        set => shape.X = value;
    }

    public virtual decimal Y
    {
        get => shape.Y;
        set => shape.Y = value;
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

    public IPresentation Presentation => shape.Presentation;
    
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

    public virtual Geometry GeometryType
    {
        get
        {
            return shape.GeometryType;
        }

        set
        {
            shape.GeometryType = value;
        }
    }

    public virtual decimal CornerSize
    {
        get
        {
            return shape.CornerSize;
        }

        set
        {
            shape.CornerSize = value;
        }
    }

    public virtual decimal[] Adjustments
    {
        get
        {
            return shape.Adjustments;
        }

        set
        {
            shape.Adjustments = value;
        }
    }

    public string? CustomData
    {
        get
        {
            return shape.CustomData;
        }

        set
        {
            shape.CustomData = value;
        }
    }

    public virtual ShapeContent ShapeContent => ShapeContent.Shape;

    public virtual IShapeOutline Outline => shape.Outline!;

    public virtual IShapeFill Fill => shape.Fill!;

    public ITextBox TextBox => textBox;

    public virtual double Rotation
    {
        get
        {
           return shape.Rotation; 
        }
    }

    public virtual bool Removable => false;

    public string SDKXPath => shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;

    public virtual ITable AsTable() => shape.AsTable();

    public virtual IMediaShape AsMedia() => shape.AsMedia();

    public void Duplicate() => shape.Duplicate();
    

    public void SetImage(string imagePath) => shape.SetImage(imagePath);

    public virtual void Remove() => shape.Remove();
    
    #endregion Composition
}