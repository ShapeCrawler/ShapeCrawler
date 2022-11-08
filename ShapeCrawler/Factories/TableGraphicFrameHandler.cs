using System;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Tables;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories;

internal class TableGraphicFrameHandler : OpenXmlElementHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

    internal override Shape? Create(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOrLayout, SCGroupShape groupShape)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var graphicData = pGraphicFrame.Graphic!.GraphicData!;
            if (!graphicData.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                return this.Successor?.Create(pShapeTreeChild, slideOrLayout, groupShape);
            }

            var table = new SlideTable(pGraphicFrame, slideOrLayout, groupShape);

            return table;
        }

        return this.Successor?.Create(pShapeTreeChild, slideOrLayout, groupShape);
    }
}