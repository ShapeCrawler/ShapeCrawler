using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.SlideShape;

internal sealed class SlideGroupShape : IGroupShape, IRemoveable
{
    private readonly P.GroupShape pGroupShape;
    private readonly SimpleShape simpleShape;

    internal SlideGroupShape(SlidePart sdkSlidePart, P.GroupShape pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.simpleShape = new SimpleShape(pGroupShape);
        this.Shapes = new SlideGroupedShapes(sdkSlidePart, pGroupShape.Elements<OpenXmlCompositeElement>());
        this.Outline = new SlideShapeOutline(sdkSlidePart, pGroupShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new SlideShapeFill(sdkSlidePart, pGroupShape.Descendants<P.ShapeProperties>().First(), false);
    }

    public IReadOnlyShapes Shapes { get; }

    #region Shape

    public int Width
    {
        get => this.simpleShape.Width();
        set => this.simpleShape.UpdateWidth(value);
    }

    public int Height
    {
        get => this.simpleShape.Height();
        set => this.simpleShape.UpdateHeight(value);
    }

    public int Id => this.simpleShape.Id();

    public string Name => this.simpleShape.Name();

    public bool Hidden => this.simpleShape.Hidden();

    public int X
    {
        get => this.simpleShape.X();
        set => this.simpleShape.UpdateX(value);
    }

    public int Y
    {
        get => this.simpleShape.Y();
        set => this.simpleShape.UpdateY(value);
    }

    #endregion Shape

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => throw new SCException(
        $"Group Shape cannot be a placeholder. Use {nameof(IShape.IsPlaceholder)} to check if the shape is a placeholder.");

    public SCGeometry GeometryType => SCGeometry.Rectangle;

    public string? CustomData
    {
        get => this.simpleShape.ParseCustomData();
        set => this.simpleShape.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.Group;
    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public IShapeFill Fill { get; }
    public bool IsTextHolder => false;

    public ITextFrame TextFrame =>
        throw new SCException(
            $"Group Shape cannot be a text holder. Use {nameof(IShape.IsTextHolder)} property to check if the shape is a text holder.");

    public double Rotation { get; }

    public ITable AsTable() =>
        throw new SCException(
            $"The Group Shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media.");

    internal void Draw(SKCanvas canvas)=>throw new System.NotImplementedException();
    internal IHtmlElement ToHtmlElement() => throw new System.NotImplementedException();
    internal string ToJson() => throw new System.NotImplementedException();
    void IRemoveable.Remove() => this.pGroupShape.Remove();
}