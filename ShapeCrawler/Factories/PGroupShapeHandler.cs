using System;
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
        private readonly GeometryFactory _geometryFactory;
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly SlidePart _slidePart;
        private readonly LocationParser _transformFactory;

        internal PGroupShapeHandler(
            ShapeContext.Builder shapeContextBuilder,
            LocationParser transformFactory,
            GeometryFactory geometryFactory,
            SlidePart sdkSldPart)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
            _slidePart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
        }

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SlideSc slide)
        {
            if (pShapeTreeChild is P.GroupShape pGroupShape)
            {
                var pShapeHandler = new AutoShapeCreator(_shapeContextBuilder, _transformFactory);
                var oleGrFrameHandler = new OleGraphicFrameHandler(_shapeContextBuilder, _transformFactory);
                var pictureHandler = new PictureHandler(_shapeContextBuilder, _transformFactory, _geometryFactory);
                var pGroupShapeHandler = new PGroupShapeHandler(_shapeContextBuilder, _transformFactory,
                    _geometryFactory, _slidePart);
                var chartGrFrameHandler = new ChartGraphicFrameHandler();
                var tableGrFrameHandler = new TableGraphicFrameHandler(_shapeContextBuilder, _transformFactory);

                pShapeHandler.Successor = pGroupShapeHandler;
                pGroupShapeHandler.Successor = oleGrFrameHandler;

                // OLE objects handler must be before pictures handler, cause OLE container can contain p:pic elements, thereby OLE as a picture
                oleGrFrameHandler.Successor = pictureHandler;
                pictureHandler.Successor = chartGrFrameHandler;
                chartGrFrameHandler.Successor = tableGrFrameHandler;

                var groupedShapes = new List<IShape>(pGroupShape.Count());
                foreach (var item in pGroupShape.OfType<OpenXmlCompositeElement>())
                {
                    var groupedShape = pShapeHandler.Create(item, slide);
                    if (groupedShape != null)
                    {
                        groupedShapes.Add(groupedShape);
                    }
                }

                var spContext = _shapeContextBuilder.Build(pShapeTreeChild);
                var transformGroup = pGroupShape.GroupShapeProperties.TransformGroup;
                var innerTransform = new NonPlaceholderTransform(transformGroup);
                var groupShape = new SlideGroupShape(innerTransform, spContext, groupedShapes, pGroupShape, slide);

                return groupShape;
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }
    }
}