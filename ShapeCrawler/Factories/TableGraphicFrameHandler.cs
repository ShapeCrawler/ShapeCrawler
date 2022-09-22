using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class TableGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

        internal override IShape Create(OpenXmlCompositeElement compositeElementOfPShapeTree, SCSlide slide, SlideGroupShape groupShape)
        {
            if (compositeElementOfPShapeTree is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData graphicData = pGraphicFrame.Graphic.GraphicData;
                if (!graphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    return this.Successor?.Create(compositeElementOfPShapeTree, slide, groupShape);
                }

                var table = new SlideTable(pGraphicFrame, slide, groupShape);

                return table;
            }

            return this.Successor?.Create(compositeElementOfPShapeTree, slide, groupShape);
        }
    }
}