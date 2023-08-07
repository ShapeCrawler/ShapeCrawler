using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services;

internal sealed class AutoShapeCreator : OpenXmlElementHandler
{
    internal override SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject,
        OneOf<ShapeCollection, SCGroupShape> shapeCollection,
        TypedOpenXmlPart slideTypedOpenXmlPart)
    {
        if (pShapeTreeChild is P.Shape pShape)
        {
            var autoShape = new SCAutoShape(pShape, slideObject, shapeCollection);
            return autoShape;
        }

        return this.Successor?.FromTreeChild(pShapeTreeChild, slideObject, shapeCollection, slideTypedOpenXmlPart);
    }
}