using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Placeholders;
using SlideDotNet.Services.ShapeCreators;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Services
{
    /// <summary>
    /// <inheritdoc cref="IShapeFactory"/>
    /// </summary>
    public class ShapeFactory : IShapeFactory
    {
        #region Fields

        private readonly IPreSettings _preSettings;
        private readonly SlidePart _sdkSldPart;

        #endregion Fields

        #region Constructors

        public ShapeFactory(SlidePart sdkSldPart, IPreSettings preSettings)
        {
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Creates collection of the shapes from sdk shape tree.
        /// </summary>
        /// <param name="sdkShapeTree"></param>
        /// <returns></returns>
        public IList<ShapeEx> FromTree(P.ShapeTree sdkShapeTree)
        {
            var sldPhFontService = new SlidePlaceholderFontService(_sdkSldPart);
            var phService = new PlaceholderService(_sdkSldPart.SlideLayoutPart);
            var transformFactory = new InnerTransformFactory(phService);
            var shapeBuilder = new ShapeEx.Builder();

            var sdkShapeHandler = new SdkShapeHandler(_preSettings, sldPhFontService, _sdkSldPart, transformFactory, shapeBuilder);
            var sdkGroupShapeHandler = new SdkGroupShapeHandler(_preSettings, sldPhFontService, _sdkSldPart, transformFactory, shapeBuilder);
            var oleGrFrameHandler = new OleGraphicFrameHandler(_preSettings, sldPhFontService, _sdkSldPart, transformFactory, shapeBuilder);
            var pictureHandler = new PictureHandler(_preSettings, sldPhFontService, _sdkSldPart, transformFactory, shapeBuilder);
            var chartGrFrameHandler = new ChartGraphicFrameHandler(_preSettings, sldPhFontService, _sdkSldPart, transformFactory, shapeBuilder);
            var tableGrFrameHandler = new TableGraphicFrameHandler(_preSettings, sldPhFontService, _sdkSldPart, transformFactory, shapeBuilder);

            sdkShapeHandler.Successor = sdkGroupShapeHandler;
            sdkGroupShapeHandler.Successor = oleGrFrameHandler;
            // OLE objects handler must be before pictures handler, cause OLE container can contain p:pic elements, thereby defining OLE as a picture
            oleGrFrameHandler.Successor = pictureHandler;
            pictureHandler.Successor = chartGrFrameHandler;
            chartGrFrameHandler.Successor = tableGrFrameHandler;

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
