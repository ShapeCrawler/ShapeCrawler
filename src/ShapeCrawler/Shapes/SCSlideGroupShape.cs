using System.Collections.Generic;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Shapes;

internal sealed class SCSlideGroupShape : IGroupShape
{
    private readonly P.GroupShape pGroupShape;
    private readonly SlidePart sdkSlidePart;
    private readonly List<ImagePart> imageParts;
    private readonly Shape shape;

    internal SCSlideGroupShape(
        P.GroupShape pGroupShape, 
        SCSlideShapes shapes,
        SlidePart sdkSlidePart,
        List<ImagePart> imageParts)
    {
        this.pGroupShape = pGroupShape;
        this.sdkSlidePart = sdkSlidePart;
        this.imageParts = imageParts;
        this.shape = new Shape(pGroupShape);
    }

    public IGroupedShapeCollection Shapes => new SlideGroupedShapes(this.pGroupShape, this, this.sdkSlidePart, this.imageParts);

    public int Width { get; set; }
    public int Height { get; set; }
    public int Id { get; }
    public string Name => this.shape.Name();
    public bool Hidden { get; }
    public IPlaceholder? Placeholder { get; }
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => SCShapeType.Group;
    public IAutoShape? AsAutoShape()
    {
        throw new System.NotImplementedException();
    }

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

    protected int GetXCoordinate()
    {
        var aXfrm = ((P.GroupShape)this.PShapeTreeChild).GroupShapeProperties!.TransformGroup!;

        return UnitConverter.HorizontalEmuToPixel(aXfrm.Offset!.X!);
    }

    protected void SetXCoordinate(int xPx)
    {
        var pGrpSpPr = this.PShapeTreeChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Offset!.X = UnitConverter.HorizontalPixelToEmu(xPx);
        aXfrm.ChildOffset!.X = UnitConverter.HorizontalPixelToEmu(xPx);
    }
    
    protected void SetYCoordinate(int yPx)
    {
        var pGrpSpPr = this.PShapeTreeChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Offset!.Y = UnitConverter.VerticalPixelToEmu(yPx);
    }
    
    protected void SetWidth(int wPx)
    {
        var pGrpSpPr = this.PShapeTreeChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Extents!.Cx = UnitConverter.HorizontalPixelToEmu(wPx);
    }
    
    protected void SetHeight(int hPx)
    {
        var pGrpSpPr = this.PShapeTreeChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Extents!.Cy = UnitConverter.VerticalPixelToEmu(hPx);
    }

    public int X { get; set; }
    public int Y { get; set; }
}