using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class PGroupShapeHandler : OpenXmlElementHandler
    {
        private readonly GeometryFactory geometryFactory;
        private readonly ShapeContext.Builder shapeContextBuilder;
        private readonly SlidePart slidePart;

        internal PGroupShapeHandler(
            ShapeContext.Builder shapeContextBuilder,
            GeometryFactory geometryFactory,
            SlidePart sdkSldPart)
        {
            this.shapeContextBuilder = shapeContextBuilder;
            this.geometryFactory = geometryFactory;
            this.slidePart = sdkSldPart;
        }

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.GroupShape pGroupShape)
            {
                var pShapeHandler = new AutoShapeCreator();
                var oleGrFrameHandler = new OleGraphicFrameHandler(this.shapeContextBuilder);
                var pictureHandler = new PictureHandler(this.shapeContextBuilder);
                var pGroupShapeHandler = new PGroupShapeHandler(this.shapeContextBuilder, this.geometryFactory, this.slidePart);
                var chartGrFrameHandler = new ChartGraphicFrameHandler();
                var tableGrFrameHandler = new TableGraphicFrameHandler(this.shapeContextBuilder);

                pShapeHandler.Successor = pGroupShapeHandler;
                pGroupShapeHandler.Successor = oleGrFrameHandler;

                // OLE objects handler must be before pictures handler, cause OLE container can contain p:pic elements, thereby OLE as a picture
                oleGrFrameHandler.Successor = pictureHandler;
                pictureHandler.Successor = chartGrFrameHandler;
                chartGrFrameHandler.Successor = tableGrFrameHandler;

                var groupedShapes = new List<IShape>(pGroupShape.Count());
                foreach (OpenXmlCompositeElement childItem in pGroupShape.OfType<OpenXmlCompositeElement>())
                {
                    IShape groupedShape = pShapeHandler.Create(childItem, slide);
                    if (groupedShape != null)
                    {
                        groupedShapes.Add(groupedShape);
                    }
                }

                var groupShape = new SlideGroupShape(groupedShapes, pGroupShape, slide);

                return groupShape;
            }

            return this.Successor?.Create(pShapeTreeChild, slide);
        }
    }
}