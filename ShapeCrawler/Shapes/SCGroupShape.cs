using OneOf;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Shapes;

internal sealed class SCGroupShape : SCShape, IGroupShape
{
    private readonly P.GroupShape pGroupShape;
    private readonly OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject;
    
    internal SCGroupShape(
        P.GroupShape pGroupShape, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pGroupShape, parentSlideObject, parentShapeCollection)
    {
        this.pGroupShape = pGroupShape;
        this.parentSlideObject = parentSlideObject;
    }

    public IGroupedShapeCollection Shapes => GroupedShapeCollection.Create(this.pGroupShape, this.parentSlideObject, this);

    public override SCShapeType ShapeType => SCShapeType.GroupShape;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
    
    protected override void SetXCoordinate(int xPx)
    {
        var pGrpSpPr = this.PShapeTreesChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Offset!.X = UnitConverter.HorizontalPixelToEmu(xPx);
    }
    
    protected override void SetYCoordinate(int yPx)
    {
        var pGrpSpPr = this.PShapeTreesChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Offset!.Y = UnitConverter.VerticalPixelToEmu(yPx);
    }
    
    protected override void SetWidth(int wPx)
    {
        var pGrpSpPr = this.PShapeTreesChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Extents!.Cx = UnitConverter.HorizontalPixelToEmu(wPx);
    }
    
    protected override void SetHeight(int hPx)
    {
        var pGrpSpPr = this.PShapeTreesChild.GetFirstChild<P.GroupShapeProperties>() !;
        var aXfrm = pGrpSpPr.TransformGroup!;
        aXfrm.Extents!.Cy = UnitConverter.VerticalPixelToEmu(hPx);
    }
}