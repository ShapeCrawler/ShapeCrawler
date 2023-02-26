using ShapeCrawler.Drawing.ShapeFill;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class SCAutoShapeFill : SCShapeFill
{
    private readonly SCAutoShape autoShape;

    internal SCAutoShapeFill(SlideStructure slideObject, P.ShapeProperties shapeProperties, SCAutoShape autoSCShape)
        : base(slideObject, shapeProperties)
    {
        this.autoShape = autoSCShape;
    }

    protected override void InitSlideBackgroundFillOr()
    {
        var pShape = (P.Shape)this.autoShape.PShapeTreeChild;
        this.useBgFill = pShape.UseBackgroundFill;
        if (this.useBgFill is not null && this.useBgFill)
        {
            this.fillType = SCFillType.SlideBackground;
        }
        else
        {
            this.aNoFill = this.properties.GetFirstChild<A.NoFill>();
            this.fillType = SCFillType.NoFill;
        }
    }
}