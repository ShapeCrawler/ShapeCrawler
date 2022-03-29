using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        internal override IShape Create(OpenXmlCompositeElement compositeElementOfPShapeTree, SCSlide slide, SlideGroupShape groupShape)
        {
            if (compositeElementOfPShapeTree is P.Shape pShape)
            {
                var autoShape = new SlideAutoShape(pShape, slide, groupShape);

                return autoShape;
            }

            return this.Successor?.Create(compositeElementOfPShapeTree, slide, groupShape);
        }
    }
}