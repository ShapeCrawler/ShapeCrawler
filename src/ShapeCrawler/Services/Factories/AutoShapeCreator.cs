using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Factories;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services.Factories;

internal sealed class AutoShapeCreator : OpenXmlElementHandler
{
    internal override SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject,
        OneOf<ShapeCollection, SCGroupShape> shapeCollection,
        ITextFrameContainer textFrameContainer)
    {
        if (pShapeTreeChild is P.Shape pShape)
        {
            var autoShape = new SCAutoShape(pShape, slideObject, shapeCollection, textFrameContainer);
            return autoShape;
        }

        return this.Successor?.FromTreeChild(pShapeTreeChild, slideObject, shapeCollection, textFrameContainer);
    }
    
}