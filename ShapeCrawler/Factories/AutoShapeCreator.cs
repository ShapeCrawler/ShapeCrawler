using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        public override IShape Create(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide)
        {
            if (pShapeTreesChild is P.Shape pShape)
            {
                var autoShape = new SlideAutoShape(slide, pShape);

                return autoShape;
            }

            return this.Successor?.Create(pShapeTreesChild, slide);
        }

        public override IShape CreateGroupedShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide, SlideGroupShape groupShape)
        {
            if (pShapeTreesChild is P.Shape pShape)
            {
                var autoShape = new SlideAutoShape(slide, pShape, groupShape);

                return autoShape;
            }

            return this.Successor?.CreateGroupedShape(pShapeTreesChild, slide, groupShape);
        }
    }
}