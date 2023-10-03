using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideShape;
using P = DocumentFormat.OpenXml.Presentation;
using Shape = ShapeCrawler.Shapes.Shape;

namespace ShapeCrawler.ShapesCollection;

internal sealed class GroupedShape : Shape
{
    private readonly P.Shape pShape;
    private readonly AutoShape autoShape;

    internal GroupedShape(TypedOpenXmlPart sdkTypedOpenXmlPart, P.Shape pShape, AutoShape autoShape)
        : base(sdkTypedOpenXmlPart, pShape)
    {
        this.pShape = pShape;
        this.autoShape = autoShape;
    }

    public override ShapeType ShapeType => this.autoShape.ShapeType;

    public override int X
    {
        get => this.autoShape.X;
        set
        {
            this.autoShape.X = value;
            var xGroupedShapePx = value; 
            var pGroupShape = this.pShape.Ancestors<P.GroupShape>().First();
            var aTransformGroup = pGroupShape.GroupShapeProperties!.TransformGroup!;
            var aOffset = aTransformGroup.Offset!;
            var aExtents = aTransformGroup.Extents!;
            var aChildOffset = aTransformGroup.ChildOffset!;
            var aChildExtents = aTransformGroup.ChildExtents!;
            var xGroupedShapeEmu = UnitConverter.HorizontalPixelToEmu(xGroupedShapePx);
            var xGroupShapeEmu = aOffset.X!;

            if (xGroupedShapeEmu < xGroupShapeEmu)
            {
                var diff = xGroupShapeEmu - xGroupedShapeEmu;
                aOffset.X = new Int64Value(aOffset.X! - diff);
                aExtents.Cx = new Int64Value(aExtents.Cx! + diff);
                aChildOffset.X = new Int64Value(aChildOffset.X! - diff);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diff);

                return;
            }

            var groupRightEmu = aOffset.X!.Value + aExtents.Cx!.Value;
            var groupedRightEmu = UnitConverter.HorizontalPixelToEmu(this.autoShape.X + this.autoShape.Width);
            if (groupedRightEmu > groupRightEmu)
            {
                var diffEmu = groupedRightEmu - groupRightEmu;
                aExtents.Cx = new Int64Value(aExtents.Cx! + diffEmu);
                aChildExtents.Cx = new Int64Value(aChildExtents.Cx! + diffEmu);
            }
        }
    }
}