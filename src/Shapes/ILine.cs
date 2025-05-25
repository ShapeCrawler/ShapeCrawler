using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a line shape.
/// </summary>
public interface ILine : IShape
{
    /// <summary>
    ///    Gets the start point of the line.
    /// </summary>
    Point StartPoint { get; }

    /// <summary>
    ///     Gets the end point of the line.
    /// </summary>
    Point EndPoint { get; }
}

internal sealed class SlideLine(Shape shape, P.ConnectionShape pConnectionShape) : ILine
{
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

    public ShapeContent ShapeContent => ShapeContent.Line;

    public IShapeOutline Outline => shape.Outline;

    public IShapeFill Fill => shape.Fill;

    public ITextBox? TextBox => shape.TextBox;

    public double Rotation => shape.Rotation;

    public string SDKXPath => shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;

    public IPresentation Presentation => shape.Presentation;

    public Geometry GeometryType
    {
        get => Geometry.Line;
        set => throw new SCException("Unable to set geometry type for line shape.");
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

    public Point StartPoint
    {
        get
        {
            var aTransform2D = pConnectionShape.GetFirstChild<P.ShapeProperties>() !.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            if (flipH && (this.Height == 0 || flipV))
            {
                return new Point(this.X, this.Y);
            }

            if (flipH)
            {
                return new Point(this.X + this.Width, this.Y);
            }

            return new Point(this.X, this.Y);
        }
    }

    public Point EndPoint
    {
        get
        {
            var aTransform2D = pConnectionShape.GetFirstChild<P.ShapeProperties>() !.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            if (this.Width == 0)
            {
                return new Point(this.X, this.Height);
            }

            if (flipH && this.Height == 0)
            {
                return new Point(this.X - this.Width, this.Y);
            }

            if (flipV)
            {
                return new Point(this.Width, this.Height);
            }

            if (flipH)
            {
                return new Point(this.X, this.Height);
            }

            return new Point(this.Width, this.Y);
        }
    }
    
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

    public bool Removable => true;

    public void Remove() => pConnectionShape.Remove();

    public ITable AsTable() => shape.AsTable();

    public IMediaShape AsMedia() => shape.AsMedia();

    public void Duplicate() => shape.Duplicate();

    public void SetText(string text) => shape.SetText(text);

    public void SetImage(string imagePath) => shape.SetImage(imagePath);

    public void SetFontName(string fontName) => shape.SetFontName(fontName);

    public void SetFontSize(decimal fontSize) => shape.SetFontSize(fontSize);

    public void SetFontColor(string colorHex) => shape.SetFontColor(colorHex);
}