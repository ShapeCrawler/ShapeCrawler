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

        public override IShape Create(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide, SlideGroupShape groupShape)
        {
            if (pShapeTreesChild is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData aGraphicData = pShapeTreesChild.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (aGraphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    SlideOLEObject oleObject = new (pGraphicFrame, slide, groupShape);

                    return oleObject;
                }
            }

            return this.Successor?.Create(pShapeTreesChild, slide, groupShape);
        }
    }
}