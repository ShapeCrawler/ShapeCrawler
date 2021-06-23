using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class TableGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";
        private readonly ShapeContext.Builder shapeContextBuilder;
        private readonly LocationParser transformFactory;

        internal TableGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory)
        {
            this.shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            this.transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
        }

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData graphicData = pGraphicFrame.Graphic.GraphicData;
                if (!graphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    return this.Successor?.Create(pShapeTreeChild, slide);
                }

                ShapeContext spContext = this.shapeContextBuilder.Build(pShapeTreeChild);
                ILocation innerTransform = this.transformFactory.FromComposite(pGraphicFrame);
                var table = new SlideTable(pGraphicFrame, innerTransform, spContext, slide);

                return table;
            }

            return this.Successor?.Create(pShapeTreeChild, slide);
        }
    }
}