using System;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services.Factories;

internal sealed class TableGraphicFrameHandler : OpenXmlElementHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

    internal override SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject,
        OneOf<ShapeCollection, SCGroupShape> shapeCollection,
        ITextFrameContainer textFrameContainer)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var graphicData = pGraphicFrame.Graphic!.GraphicData!;
            if (!graphicData.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
            {
                return this.Successor?.FromTreeChild(pShapeTreeChild, slideObject, shapeCollection, textFrameContainer);
            }

            var table = new SCTable(pGraphicFrame, slideObject, shapeCollection);

            return table;
        }

        return this.Successor?.FromTreeChild(pShapeTreeChild, slideObject, shapeCollection, textFrameContainer);
    }
}