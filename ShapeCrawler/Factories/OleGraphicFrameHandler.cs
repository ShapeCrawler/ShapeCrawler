using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using System;
using ShapeCrawler.OLEObjects;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class OleGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";

        internal override Shape Create(OpenXmlCompositeElement compositeElementOfPShapeTree, SCSlide slide, SlideGroupShape groupShape)
        {
            if (compositeElementOfPShapeTree is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData aGraphicData = compositeElementOfPShapeTree.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (aGraphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    SlideOLEObject oleObject = new (pGraphicFrame, slide, groupShape);

                    return oleObject;
                }
            }

            return this.Successor?.Create(compositeElementOfPShapeTree, slide, groupShape);
        }
    }
}