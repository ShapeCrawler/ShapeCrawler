using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Services.Placeholders;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Services.ShapeCreators
{
    /// <summary>
    /// <inheritdoc cref="IShapeFactory"/>
    /// </summary>
    public class ShapeFactory : IShapeFactory
    {
        #region Fields

        private readonly IPreSettings _preSettings;

        #endregion Fields

        #region Constructors

        public ShapeFactory(IPreSettings preSettings)
        {
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
        }

        #endregion Constructors

        #region Public Methods

        public IList<ShapeEx> FromSldPart(SlidePart sdkSldPart)
        {
            var sldPhFontService = new PlaceholderFontService(sdkSldPart); //TODO: make DI
            var phService = new PlaceholderService(sdkSldPart.SlideLayoutPart);
            var transformFactory = new LocationParser(phService);
            var geometryFactory = new GeometryFactory(phService);
            var shapeContextBuilder = new ShapeContext.Builder(_preSettings, sldPhFontService,sdkSldPart);

            var chartGrFrameHandler = new ChartGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var tableGrFrameHandler = new TableGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var oleGrFrameHandler = new OleGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var sdkShapeHandler = new SdkShapeHandler(shapeContextBuilder, transformFactory, geometryFactory);
            var pictureHandler = new PictureHandler(shapeContextBuilder, transformFactory, geometryFactory, sdkSldPart);
            var sdkGroupShapeHandler = new SdkGroupShapeHandler(shapeContextBuilder, transformFactory, geometryFactory, sdkSldPart);

            sdkShapeHandler.Successor = sdkGroupShapeHandler;
            sdkGroupShapeHandler.Successor = oleGrFrameHandler;
            // OLE objects handler must be before pictures handler, cause OLE container can contain p:pic elements,
            // thereby defining OLE as a picture
            oleGrFrameHandler.Successor = pictureHandler;
            pictureHandler.Successor = chartGrFrameHandler;
            chartGrFrameHandler.Successor = tableGrFrameHandler;

            var sdkShapeTree = sdkSldPart.Slide.CommonSlideData.ShapeTree;
            var shapes = new List<ShapeEx>(sdkShapeTree.Count());
            foreach (var openXmlElement in sdkShapeTree)
            {
                var shape = sdkShapeHandler.Create(openXmlElement);
                if (shape != null)
                {
                    shapes.Add(shape);
                }
            }

            return shapes;
        }

        #endregion Public Methods
    }
}
