using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Models.Transforms;
using SlideDotNet.Services.Builders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services.ShapeCreators
{
    public class SdkGroupShapeHandler : OpenXmlElementHandler
    {
        private readonly IPreSettings _preSettings;
        private readonly SlidePlaceholderFontService _sldPhFontService;
        private readonly SlidePart _sdkSldPart;
        private readonly InnerTransformFactory _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;

        public SdkGroupShapeHandler(IPreSettings preSettings,
            SlidePlaceholderFontService sldFontService,
            SlidePart sdkSldPart,
            InnerTransformFactory transformFactory,
            IShapeBuilder shapeBuilder)
        {
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _sldPhFontService = sldFontService ?? throw new ArgumentNullException(nameof(sldFontService));
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _shapeBuilder = shapeBuilder;
        }

        public override ShapeEx Create(OpenXmlElement openXmlElement)
        {
            Check.NotNull(openXmlElement, nameof(openXmlElement));

            if (openXmlElement is P.GroupShape sdkGroupShape)
            {
                var sdkShapeHandler = new SdkShapeHandler(_preSettings, _sldPhFontService, _sdkSldPart, _transformFactory, _shapeBuilder);
                var sdkGroupShapeHandler = new SdkGroupShapeHandler(_preSettings, _sldPhFontService, _sdkSldPart, _transformFactory, _shapeBuilder);
                var oleGrFrameHandler = new OleGraphicFrameHandler(_preSettings, _sldPhFontService, _sdkSldPart, _transformFactory, _shapeBuilder);
                var pictureHandler = new PictureHandler(_preSettings, _sldPhFontService, _sdkSldPart, _transformFactory, _shapeBuilder);
                var chartGrFrameHandler = new ChartGraphicFrameHandler(_preSettings, _sldPhFontService, _sdkSldPart, _transformFactory, _shapeBuilder);
                var tableGrFrameHandler = new TableGraphicFrameHandler(_preSettings, _sldPhFontService, _sdkSldPart, _transformFactory, _shapeBuilder);

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
                var spContext = new ShapeContext(_preSettings, _sldPhFontService, sdkGroupShape, _sdkSldPart);
                var transformGroup = sdkGroupShape.GroupShapeProperties.TransformGroup;
                var innerTransform = new NonPlaceholderTransform(transformGroup); //TODO: use factory version instead
                var shape = _shapeBuilder.WithGroup(innerTransform, spContext, groupedShapes);

                return shape;
            }

            if (Successor != null)
            {
                return Successor.Create(openXmlElement);
            }

            return null;
        }
    }
}