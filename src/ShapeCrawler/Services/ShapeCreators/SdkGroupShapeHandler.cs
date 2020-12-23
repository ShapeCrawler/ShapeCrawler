using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.Transforms;
using ShapeCrawler.Services.Builders;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services.ShapeCreators
{
    public class SdkGroupShapeHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly SlidePart _sdkSldPart;
        private readonly IGeometryFactory _geometryFactory; //TODO: DI
        private readonly LocationParser _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;

        public SdkGroupShapeHandler(ShapeContext.Builder shapeContextBuilder,
                                    LocationParser transformFactory,
                                    IGeometryFactory geometryFactory,
                                    SlidePart sdkSldPart) :
            this(shapeContextBuilder, transformFactory, geometryFactory, sdkSldPart, new ShapeEx.Builder())
        {
 
        }

        public SdkGroupShapeHandler(ShapeContext.Builder shapeContextBuilder,
                                    LocationParser transformFactory,
                                    IGeometryFactory geometryFactory,
                                    SlidePart sdkSldPart,
                                    IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
        }

        public override ShapeEx Create(OpenXmlElement sdkElement)
        {
            Check.NotNull(sdkElement, nameof(sdkElement));

            if (sdkElement is P.GroupShape sdkGroupShape)
            {
                var sdkShapeHandler = new SdkShapeHandler(_shapeContextBuilder, _transformFactory, _geometryFactory);
                var oleGrFrameHandler = new OleGraphicFrameHandler(_shapeContextBuilder, _transformFactory, _shapeBuilder);
                var pictureHandler = new PictureHandler(_shapeContextBuilder, _transformFactory, _geometryFactory, _sdkSldPart);
                var sdkGroupShapeHandler = new SdkGroupShapeHandler(_shapeContextBuilder, _transformFactory, _geometryFactory, _sdkSldPart);
                var chartGrFrameHandler = new ChartGraphicFrameHandler(_shapeContextBuilder, _transformFactory, _shapeBuilder);
                var tableGrFrameHandler = new TableGraphicFrameHandler(_shapeContextBuilder, _transformFactory, _shapeBuilder);

                sdkShapeHandler.Successor = sdkGroupShapeHandler;
                sdkGroupShapeHandler.Successor = oleGrFrameHandler;
                // OLE objects handler must be before pictures handler, cause OLE container can contain p:pic elements, thereby OLE as a picture
                oleGrFrameHandler.Successor = pictureHandler;
                pictureHandler.Successor = chartGrFrameHandler;
                chartGrFrameHandler.Successor = tableGrFrameHandler;

                var groupedShapes = new List<ShapeEx>(sdkGroupShape.Count());
                foreach (var item in sdkGroupShape)
                {
                    var groupedShape = sdkShapeHandler.Create(item);
                    if (groupedShape != null)
                    {
                        groupedShapes.Add(groupedShape);
                    }
                }
                var spContext = _shapeContextBuilder.Build(sdkElement);
                var transformGroup = sdkGroupShape.GroupShapeProperties.TransformGroup;
                var innerTransform = new NonPlaceholderTransform(transformGroup); //TODO: use factory version instead
                var shape = _shapeBuilder.WithGroup(innerTransform, spContext, groupedShapes);

                return shape;
            }

            if (Successor != null)
            {
                return Successor.Create(sdkElement);
            }

            return null;
        }
    }
}