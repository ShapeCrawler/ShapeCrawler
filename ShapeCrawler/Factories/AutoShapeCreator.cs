using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.Shape pShape)
            {
                var autoShape = new SlideAutoShape(pShape, slide);

                return autoShape;
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }
    }
}