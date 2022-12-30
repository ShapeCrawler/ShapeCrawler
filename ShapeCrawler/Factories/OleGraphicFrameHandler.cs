using System;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories;

internal sealed class OleGraphicFrameHandler : OpenXmlElementHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";

    internal override Shape? Create(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, SCGroupShape groupShape)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var aGraphicData = pShapeTreeChild!.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>();
            if (aGraphicData!.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                 var oleObject = new SlideOLEObject (pGraphicFrame, slideObject, groupShape);

                 return oleObject;
            }
        }

        return this.Successor?.Create(pShapeTreeChild, slideObject, groupShape);
    }
}