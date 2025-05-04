using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Groups;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a group shape.
/// </summary>
public interface IGroup : IShape
{
    /// <summary>
    ///     Gets grouped shape collection.
    /// </summary>
    IShapeCollection Shapes { get; }
}

internal sealed class Group : IGroup
{
    private readonly P.GroupShape pGroupShape;
    private readonly Shape shape;

    internal Group(Shape shape, P.GroupShape pGroupShape)
    {
        this.shape = shape;
        this.pGroupShape = pGroupShape;
        this.Shapes = new GroupedShapeCollection(pGroupShape.Elements<OpenXmlCompositeElement>());
        var pShapeProperties = pGroupShape.Descendants<P.ShapeProperties>().First();
        this.Outline = new SlideShapeOutline(pShapeProperties);
        this.Fill = new ShapeFill(pShapeProperties);
    }

    public IShapeCollection Shapes { get; }

    public Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("The geometry type of a group shape cannot be set.");
    }

    public PlaceholderType? PlaceholderType => null;

    public ShapeContent ShapeContent => ShapeContent.Group;

    public bool HasOutline => true;

    public IShapeOutline Outline { get; }

    public bool HasFill => true;

    public IShapeFill Fill { get; }

    public ITextBox? TextBox => null;

    public bool Removable => true;

    public decimal CornerSize
    {
        get => this.shape.CornerSize;
        set => this.shape.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => this.shape.Adjustments;
        set => this.shape.Adjustments = value;
    }

    public decimal Width
    {
        get => this.shape.Width;
        set => this.shape.Width = value;
    }

    public decimal Height
    {
        get => this.shape.Height;
        set => this.shape.Height = value;
    }

    public int Id => this.shape.Id;

    public string Name
    {
        get => this.shape.Name;
        set => this.shape.Name = value;
    }

    public string AltText
    {
        get => this.shape.AltText;
        set => this.shape.AltText = value;
    }

    public bool Hidden => this.shape.Hidden;

    public string? CustomData
    {
        get => this.shape.CustomData;
        set => this.shape.CustomData = value;
    }

    public decimal X
    {
        get => this.shape.X;
        set => this.shape.X = value;
    }

    public decimal Y
    {
        get => this.shape.Y;
        set => this.shape.Y = value;
    }

    public double Rotation
    {
        get
        {
            var aTransformGroup = this.pGroupShape.GroupShapeProperties!.TransformGroup!;
            var rotation = aTransformGroup.Rotation?.Value ?? 0;
            return rotation / 60000d;
        }
    }

    public string SDKXPath => this.shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => this.shape.SDKOpenXmlElement;

    public IPresentation Presentation => this.shape.Presentation;
    
    public void Duplicate() => this.shape.Duplicate();
    
    public IShape Shape(string shape) => this.Shapes.Shape(shape);

    public void Remove() => this.shape.Remove();

    public ITable AsTable()
    {
        throw new System.NotImplementedException();
    }

    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is a media (audio, video, etc.");
}