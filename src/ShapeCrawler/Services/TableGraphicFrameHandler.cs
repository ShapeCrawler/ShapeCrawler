using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services;

internal sealed class TableGraphicFrameHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

    internal SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ChartWorkbook> chartWorkbooks)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var graphicData = pGraphicFrame.Graphic!.GraphicData!;
            if (!graphicData.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                return this.Successor?.FromTreeChild(pShapeTreeChild, slideOf, shapeCollectionOf, slideTypedOpenXmlPart, chartWorkbooks);
            }

            var table = new SCTable(pGraphicFrame, slideOf, shapeCollectionOf, slideTypedOpenXmlPart);

            return table;
        }

        return this.Successor?.FromTreeChild(pShapeTreeChild, slideOf, shapeCollectionOf, slideTypedOpenXmlPart);
    }
}