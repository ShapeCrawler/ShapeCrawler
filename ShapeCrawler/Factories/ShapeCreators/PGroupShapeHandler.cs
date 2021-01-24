using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories.Builders;
using ShapeCrawler.Models.Transforms;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories.ShapeCreators
{
    public class PGroupShapeHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly SlidePart _sdkSldPart;
        private readonly GeometryFactory _geometryFactory;
        private readonly LocationParser _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;

        public PGroupShapeHandler(ShapeContext.Builder shapeContextBuilder,
                                    LocationParser transformFactory,
                                    GeometryFactory geometryFactory,
                                    SlidePart sdkSldPart) :
            this(shapeContextBuilder, transformFactory, geometryFactory, sdkSldPart, new ShapeSc.Builder())
        {
 
        }

        public PGroupShapeHandler(ShapeContext.Builder shapeContextBuilder,
                                    LocationParser transformFactory,
                                    GeometryFactory geometryFactory,
                                    SlidePart sdkSldPart,
                                    IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
        }

        public override ShapeSc Create(OpenXmlCompositeElement shapeTreeSource)
        {
            Check.NotNull(shapeTreeSource, nameof(shapeTreeSource));

            if (shapeTreeSource is P.GroupShape pGroupShape)
            {
                var pShapeHandler = new PShapeHandler(_shapeContextBuilder, _transformFactory, _geometryFactory);
                var oleGrFrameHandler = new OleGraphicFrameHandler(_shapeContextBuilder, _transformFactory, _shapeBuilder);
                var pictureHandler = new PictureHandler(_shapeContextBuilder, _transformFactory, _geometryFactory, _sdkSldPart);
                var pGroupShapeHandler = new PGroupShapeHandler(_shapeContextBuilder, _transformFactory, _geometryFactory, _sdkSldPart);
                var chartGrFrameHandler = new ChartGraphicFrameHandler(_shapeContextBuilder, _transformFactory, _shapeBuilder);
                var tableGrFrameHandler = new TableGraphicFrameHandler(_shapeContextBuilder, _transformFactory, _shapeBuilder);

                pShapeHandler.Successor = pGroupShapeHandler;
                pGroupShapeHandler.Successor = oleGrFrameHandler;
                // OLE objects handler must be before pictures handler, cause OLE container can contain p:pic elements, thereby OLE as a picture
                oleGrFrameHandler.Successor = pictureHandler;
                pictureHandler.Successor = chartGrFrameHandler;
                chartGrFrameHandler.Successor = tableGrFrameHandler;

                var groupedShapes = new List<ShapeSc>(pGroupShape.Count());
                foreach (var item in pGroupShape.OfType<OpenXmlCompositeElement>())
                {
                    var groupedShape = pShapeHandler.Create(item);
                    if (groupedShape != null)
                    {
                        groupedShapes.Add(groupedShape);
                    }
                }
                var spContext = _shapeContextBuilder.Build(shapeTreeSource);
                var transformGroup = pGroupShape.GroupShapeProperties.TransformGroup;
                var innerTransform = new NonPlaceholderTransform(transformGroup);
                var shape = _shapeBuilder.WithGroup(innerTransform, spContext, groupedShapes, pGroupShape);

                return shape;
            }

            if (Successor != null)
            {
                return Successor.Create(shapeTreeSource);
            }

            return null;
        }
    }
}