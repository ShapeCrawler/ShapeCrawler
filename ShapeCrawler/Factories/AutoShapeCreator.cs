using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories;

internal sealed class AutoShapeCreator : OpenXmlElementHandler
{
    internal override SCShape? Create(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject,
        OneOf<ShapeCollection, SCGroupShape> shapeCollection)
    {
        if (pShapeTreeChild is P.Shape pShape)
        {
            var autoShape = new SCAutoShape(pShape, slideObject, shapeCollection);
            return autoShape;
        }

        return this.Successor?.Create(pShapeTreeChild, slideObject, shapeCollection);
    }
}