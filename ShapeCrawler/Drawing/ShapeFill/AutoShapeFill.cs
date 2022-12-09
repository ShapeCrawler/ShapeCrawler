using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing.ShapeFill;

internal class AutoShapeFill : ShapeFill
{
    private readonly AutoShape autoShape;

    internal AutoShapeFill(SlideObject slideObject, P.ShapeProperties shapeProperties, AutoShape autoShape)
        : base(slideObject, shapeProperties)
    {
        this.autoShape = autoShape;
    }

    protected override void InitSlideBackgroundFillOr()
    {
        var pShape = (P.Shape)this.autoShape.PShapeTreesChild;
        this.useBgFill = pShape.UseBackgroundFill;
        if (this.useBgFill is not null && this.useBgFill)
        {
            this.fillType = SCFillType.SlideBackground;
        }
        else
        {
            this.aNoFill = this.framePr.GetFirstChild<A.NoFill>();
            this.fillType = SCFillType.NoFill;
        }
    }
}