using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing.ShapeFill;

internal sealed class SCAutoSCShapeFill : SCShapeFill
{
    private readonly AutoSCShape _autoSCShape;

    internal SCAutoSCShapeFill(SlideObject slideObject, P.ShapeProperties shapeProperties, AutoSCShape autoSCShape)
        : base(slideObject, shapeProperties)
    {
        this._autoSCShape = autoSCShape;
    }

    protected override void InitSlideBackgroundFillOr()
    {
        var pShape = (P.Shape)this._autoSCShape.PShapeTreesChild;
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