using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services;

internal sealed class OleGraphicFrameHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";

    internal SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ChartWorkbook> chartWorkbooks)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var aGraphicData = pShapeTreeChild!.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>();
            if (aGraphicData!.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                var oleObject = new SCOLEObject(pGraphicFrame, slideOf, shapeCollectionOf);

                return oleObject;
            }
        }

        return this.Successor?.FromTreeChild(pShapeTreeChild, slideOf, shapeCollectionOf, slideTypedOpenXmlPart);
    }
}