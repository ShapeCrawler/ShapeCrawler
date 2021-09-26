using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using System;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class OleGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";
        private readonly ShapeContext.Builder shapeContextBuilder;

        internal OleGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder)
        {
            this.shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
        }

        public override IShape Create(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide, SlideGroupShape groupShape)
        {
            if (pShapeTreesChild is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData aGraphicData = pShapeTreesChild.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (aGraphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    ShapeContext spContext = this.shapeContextBuilder.Build(pShapeTreesChild);
                    SlideOLEObject oleObject = new (pGraphicFrame, slide, spContext, groupShape);

                    return oleObject;
                }
            }

            return this.Successor?.Create(pShapeTreesChild, slide, groupShape);
        }
    }
}