using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        internal override Shape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide, SlideGroupShape groupShape)
        {
            if (pShapeTreeChild is P.Shape pShape)
            {
                var slideAutoShape = new SlideAutoShape(pShape, slide, groupShape);
                return slideAutoShape;
            }

            return this.Successor?.Create(pShapeTreeChild, slide, groupShape);
        }
    }
}