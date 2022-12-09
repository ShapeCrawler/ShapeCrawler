using System;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories;

internal class TableGraphicFrameHandler : OpenXmlElementHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

    internal override Shape? Create(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, SCGroupShape groupShape)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var graphicData = pGraphicFrame.Graphic!.GraphicData!;
            if (!graphicData.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                return this.Successor?.Create(pShapeTreeChild, slideObject, groupShape);
            }

            var table = new SCTable(pGraphicFrame, slideObject, groupShape);

            return table;
        }

        return this.Successor?.Create(pShapeTreeChild, slideObject, groupShape);
    }
}