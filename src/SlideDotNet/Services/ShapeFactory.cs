using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Extensions;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Models.SlideComponents.Chart;
using SlideDotNet.Models.Transforms;
using SlideDotNet.Services.Builders;
using SlideDotNet.Services.Placeholders;
using SlideXML.Models.SlideComponents;
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

        private readonly IShapeBuilder _shapeBuilder;
        private readonly IPreSettings _preSettings;
        private readonly IPlaceholderService _phService;
        private readonly SlidePlaceholderFontService _slideFontService;
        private readonly SlidePart _sdkSldPart;
        private P.GroupShapeType _currentSdkGroupType;

        #endregion Fields

        #region Constructors

        public ShapeFactory(SlidePart xmlSldPart, IPreSettings preSettings)
        {
            _sdkSldPart = xmlSldPart ?? throw new ArgumentNullException(nameof(xmlSldPart));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _slideFontService = new SlidePlaceholderFontService(xmlSldPart);
            _phService = new PlaceholderService(xmlSldPart.SlideLayoutPart);
            _shapeBuilder = new ShapeEx.Builder();
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
            _currentSdkGroupType = sdkShapeTree ?? throw new ArgumentNullException(nameof(sdkShapeTree));
           
            var skdTreeElements = _currentSdkGroupType.Elements<OpenXmlCompositeElement>();
            var treeShapes = FromElements(skdTreeElements);

            var sdkGroupShapes = skdTreeElements.OfType<P.GroupShape>();
            AddGroups(treeShapes, sdkGroupShapes);

            return treeShapes;
        }

        #endregion Public Methods

        #region Private Methods

        private IEnumerable<ShapeEx> FromGroup(P.GroupShape sdkGroupShape)
        {
            _currentSdkGroupType = sdkGroupShape;

            var sdkGroupedElements = sdkGroupShape.Elements<OpenXmlCompositeElement>();
            var groupedShapes = FromElements(sdkGroupedElements);

            return groupedShapes;
        }

        private List<ShapeEx> FromElements(IEnumerable<OpenXmlCompositeElement> allXmlElements)
        {
            var shapesCollection = new List<ShapeEx>(allXmlElements.Count());
            var xmlGraphicFrameElements = allXmlElements.OfType<P.GraphicFrame>();

            // OLE objects
            var xmlOleGraphicFrames = xmlGraphicFrameElements.Where(e => e.Descendants<P.OleObject>().Any());
            AddOleObjects(shapesCollection, xmlOleGraphicFrames);

            // Pictures
            var xmlPictures = xmlGraphicFrameElements.Except(xmlOleGraphicFrames).SelectMany(e => e.Descendants<P.Picture>());
            var allXmlPicElements = allXmlElements.OfType<P.Picture>().Union(xmlPictures);
            AddPictures(shapesCollection, allXmlPicElements);

            // AutoShapes
            var xmlShapes = allXmlElements.OfType<P.Shape>();
            AddAutoShapes(shapesCollection, xmlShapes);

            // Tables
            var xmlTableGraphicFrames = xmlGraphicFrameElements.Where(g => g.Descendants<A.Table>().Any());
            AddTables(shapesCollection, xmlTableGraphicFrames);

            // Charts
            var xmlChartGraphicFrames = xmlGraphicFrameElements.Where(g => g.HasChart());
            AddCharts(shapesCollection, xmlChartGraphicFrames);

            return shapesCollection;
        }

        private void AddOleObjects(List<ShapeEx> shapesCollection, IEnumerable<P.GraphicFrame> graphicFramesOle)
        {
            foreach (var xmlGraphicFrame in graphicFramesOle)
            {
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGraphicFrame, _sdkSldPart);
                var innerTransform = GetTransform(xmlGraphicFrame);

                var ole = new OleObject(xmlGraphicFrame);
                var newShape = _shapeBuilder.WithOle(innerTransform, spContext, ole);

                shapesCollection.Add(newShape);
            }
        }

        private void AddTables(List<ShapeEx> shapesCollection, IEnumerable<P.GraphicFrame> graphicFramesTable)
        {
            foreach (var xmlGraphicFrame in graphicFramesTable)
            {
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGraphicFrame, _sdkSldPart);
                var innerTransform = GetTransform(xmlGraphicFrame);

                var table = new TableEx(xmlGraphicFrame, spContext);
                var newShape = _shapeBuilder.WithTable(innerTransform, spContext, table);

                shapesCollection.Add(newShape);
            }
        }

        private void AddPictures(List<ShapeEx> shapesCollection, IEnumerable<P.Picture> xmlPictures)
        {
            foreach (var xmlPic in xmlPictures)
            {
                // Picture
                var pBlipFill = xmlPic.GetFirstChild<P.BlipFill>();
                var blipRelateId = pBlipFill?.Blip?.Embed?.Value;
                if (blipRelateId == null)
                {
                    continue;
                }
                var picture = new Picture(_sdkSldPart, blipRelateId);

                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlPic, _sdkSldPart);

                var innerTransform = GetTransform(xmlPic);

                var newPicShape = _shapeBuilder.WithPicture(innerTransform, spContext, picture);
                shapesCollection.Add(newPicShape);
            }
        }

        private void AddAutoShapes(List<ShapeEx> shapesCollection, IEnumerable<P.Shape> xmlShapes)
        {
            foreach (var xmlShape in xmlShapes)
            {
                ShapeEx newAutoShape;
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlShape, _sdkSldPart);
                var t2d = xmlShape.ShapeProperties.Transform2D;
                IInnerTransform innerTransform;
                if (t2d != null)
                {
                    innerTransform = GetTransform(t2d);

                    newAutoShape = _shapeBuilder.WithAutoShape(innerTransform, spContext);
                }
                else // is placeholder obviously
                {
                    var placeholderLocationData = _phService.TryGet(xmlShape);
                    innerTransform = new PlaceholderTransform(placeholderLocationData);
                    newAutoShape = _shapeBuilder.WithAutoShape(innerTransform, spContext);
                }

                shapesCollection.Add(newAutoShape);
            }
        }

        private void AddCharts(List<ShapeEx> shapesCollection, IEnumerable<P.GraphicFrame> xmlTablesGraphicFrames)
        {
            foreach (var xmlGraphicFrame in xmlTablesGraphicFrames)
            {
                var chartEx = new ChartEx(xmlGraphicFrame, _sdkSldPart);
                var innerTransform = GetTransform(xmlGraphicFrame);
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGraphicFrame, _sdkSldPart);
                var newShape = _shapeBuilder.WithChart(innerTransform, spContext, chartEx);
                shapesCollection.Add(newShape);
            }
        }

        private void AddGroups(List<ShapeEx> shapesCollection, IEnumerable<P.GroupShape> sdkGroupShapes)
        {
            foreach (var xmlGroup in sdkGroupShapes)
            {
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGroup, _sdkSldPart);
                var groupedShapes = FromGroup(xmlGroup);
                var transformGroup = xmlGroup.GroupShapeProperties.TransformGroup;
                var innerTransform = new NonPlaceholderTransform(transformGroup);

                var newShape = _shapeBuilder.WithGroup(innerTransform, spContext, groupedShapes);
                shapesCollection.Add(newShape);
            }
        }

        private IInnerTransform GetTransform(OpenXmlCompositeElement sdkElement)
        {
            IInnerTransform innerTransform; //TODO: move get IInnerTransform methods into separate class
            if (_currentSdkGroupType is P.GroupShape groupShape)
            {
                innerTransform = new NonPlaceholderGroupedTransform(sdkElement, groupShape);
            }
            else
            {
                innerTransform = new NonPlaceholderTransform(sdkElement);
            }

            return innerTransform;
        }

        private IInnerTransform GetTransform(P.Picture xmlPic)
        {
            IInnerTransform innerTransform;
            var t2d = xmlPic.ShapeProperties.Transform2D;
            if (t2d != null)
            {
                innerTransform = GetTransform(t2d);
            }
            else
            {
                var placeholderLocationData = _phService.TryGet(xmlPic);
                innerTransform = new PlaceholderTransform(placeholderLocationData);
            }

            return innerTransform;
        }

        #endregion Private Methods
    }
}
