using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories;

internal sealed class AutoShapeCreator : OpenXmlElementHandler
{
    internal override Shape? Create(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, SCGroupShape groupShape)
    {
        if (pShapeTreeChild is P.Shape pShape)
        {
            var slideAutoShape = new AutoShape(pShape, slideObject, groupShape);
            return slideAutoShape;
        }

        return this.Successor?.Create(pShapeTreeChild, slideObject, groupShape);
    }
}