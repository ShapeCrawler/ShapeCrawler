using System.Collections.Generic;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Shapes;

internal sealed record SlideGroupShape : IGroupShape, IRemoveable
{
    private readonly P.GroupShape pGroupShape;
    private readonly Shape shape;

    internal SlideGroupShape(SlidePart sdkSlidePart, P.GroupShape pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.shape = new Shape(pGroupShape);
        this.Shapes = new SlideGroupedShapes(sdkSlidePart, pGroupShape.Elements<OpenXmlCompositeElement>());
    }

    internal void OnGroupedShapeXChanged(object? sender, int xGroupedShape)
    {
        var offset = this.ATransformGroup.Offset!;
        var extents = this.ATransformGroup.Extents!;
        var childOffset = this.ATransformGroup.ChildOffset!;
        var childExtents = this.ATransformGroup.ChildExtents!;

        if (xGroupedShape < this.X)
        {
            var groupedXEmu = UnitConverter.HorizontalPixelToEmu(xGroupedShape);
            var diff = this.ATransformGroup.Offset!.X! - groupedXEmu;

            offset.X = new Int64Value(offset.X! - diff);
            extents.Cx = new Int64Value(extents.Cx! + diff);
            childOffset.X = new Int64Value(childOffset.X! - diff);
            childExtents.Cx = new Int64Value(childExtents.Cx! + diff);

            return;
        }

        var groupedShape = (IShape)sender!;
        var parentGroupRight = this.X + this.Width;
        var groupedShapeRight = groupedShape.X + groupedShape.Width;
        if (groupedShapeRight > parentGroupRight)
        {
            var diff = groupedShapeRight - parentGroupRight;
            var diffEmu = UnitConverter.HorizontalPixelToEmu(diff);
            extents.Cx = new Int64Value(extents.Cx! + diffEmu);
            childExtents.Cx = new Int64Value(childExtents.Cx! + diffEmu);
        }
    }

    internal void OnGroupedShapeYChanged(object? sender, int yGroupedShape)
    {
        var offset = this.ATransformGroup.Offset!;
        var extents = this.ATransformGroup.Extents!;
        var childOffset = this.ATransformGroup.ChildOffset!;
        var childExtents = this.ATransformGroup.ChildExtents!;

        if (yGroupedShape < this.Y)
        {
            var groupedYEmu = UnitConverter.VerticalPixelToEmu(yGroupedShape);
            var diff = this.ATransformGroup.Offset!.Y! - groupedYEmu;

            offset.Y = new Int64Value(offset.Y! - diff);
            extents.Cy = new Int64Value(extents.Cy! + diff);
            childOffset.Y = new Int64Value(childOffset.Y! - diff);
            childExtents.Cy = new Int64Value(childExtents.Cy! + diff);

            return;
        }

        var groupedShape = (IShape)sender!;
        var parentGroupBottom = this.Y + this.Height;
        var groupedShapeBottom = groupedShape.Y + groupedShape.Height;
        if (groupedShapeBottom > parentGroupBottom)
        {
            var diff = groupedShapeBottom - parentGroupBottom;
            var diffEmu = UnitConverter.HorizontalPixelToEmu(diff);
            extents.Cy = new Int64Value(extents.Cy! + diffEmu);
            childExtents.Cy = new Int64Value(childExtents.Cy! + diffEmu);
        }
    }

    public IReadOnlyShapeCollection Shapes { get; }

    public int Width
    {
        get => this.shape.Width();
        set => this.shape.UpdateWidth(value);
    }

    public int Height
    {
        get => this.shape.Height();
        set => this.shape.UpdateHeight(value);
    }

    public int Id => this.shape.Id();

    public string Name => this.shape.Name();

    public bool Hidden => this.shape.Hidden();
    public bool IsPlaceholder() => false;

    public IPlaceholder? Placeholder { get; }
    public SCGeometry GeometryType { get; }

    public string? CustomData
    {
        get => this.shape.CustomData();
        set => this.shape.UpdateCustomData(value);
    }

    public SCShapeType ShapeType => SCShapeType.Group;

    public IAutoShape AsAutoShape() =>
        throw new SCException(
            $"The shape is not an AutoShape. Use {nameof(IGroupShape.ShapeType)} method to check the shape type.");

    internal A.TransformGroup ATransformGroup => this.pGroupShape.GroupShapeProperties!.TransformGroup!;

    internal void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new System.NotImplementedException();
    }

    internal string ToJson()
    {
        throw new System.NotImplementedException();
    }

    public int X
    {
        get => this.shape.X();
        set => this.shape.UpdateX(value);
    }

    public int Y
    {
        get => this.shape.Y();
        set => this.shape.UpdateY(value);
    }

    void IRemoveable.Remove()
    {
        this.pGroupShape.Remove();
    }
}