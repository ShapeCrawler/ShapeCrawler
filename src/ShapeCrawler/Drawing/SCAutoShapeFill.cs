using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class SCAutoShapeFill : SCShapeFill
{
    private readonly SCSlideAutoShape autoShape;

    internal SCAutoShapeFill(
        ISlideStructure slideStructure, 
        P.ShapeProperties shapeProperties, 
        SCSlideAutoShape autoShape, 
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ImagePart> imageParts)
        : base(slideStructure, shapeProperties, slideTypedOpenXmlPart, imageParts)
    {
        this.autoShape = autoShape;
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
            this.fillType = SCFillType.NoFill;
        }
    }
}