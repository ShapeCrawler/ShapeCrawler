using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class ChartGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SlideSc slide)
        {
            if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (aGraphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    SlideChart chart = new (pGraphicFrame, slide);

                    return chart;
                }
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }
    }
}