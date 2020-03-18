using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Extensions;
using SlideDotNet.Models;
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

        private readonly IShapeBuilder _shapeBuilder = new ShapeEx.Builder(); // TODO: [DI]
        private readonly IPreSettings _preSettings;
        private readonly IPlaceholderService _phService;
        private readonly SlidePlaceholderFontService _slideFontService;
        private readonly SlidePart _xmlSldPart;
        private P.GroupShapeType _currentSdkGroupType;

        #endregion Fields

        #region Constructors

        public ShapeFactory(SlidePart xmlSldPart, IPreSettings preSettings)
        {
            _xmlSldPart = xmlSldPart ?? throw new ArgumentNullException(nameof(xmlSldPart));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _slideFontService = new SlidePlaceholderFontService(xmlSldPart);
            _phService = new PlaceholderService(xmlSldPart.SlideLayoutPart);
        }

        #endregion Constructors

        #region Public Methods

        public IList<ShapeEx> FromTree(P.ShapeTree sdkShapeTree)
        {
            _currentSdkGroupType = sdkShapeTree ?? throw new ArgumentNullException(nameof(sdkShapeTree));
           
            var allXmlElements = _currentSdkGroupType.Elements<OpenXmlCompositeElement>();
            var shapesCollection = new List<ShapeEx>(allXmlElements.Count());
            ParseElements(shapesCollection, allXmlElements);

            var xmlGroups = allXmlElements.OfType<P.GroupShape>();
            AddGroups(shapesCollection, xmlGroups);

            return shapesCollection;
        }

        #endregion Public Methods

        #region Private Methods

        private IList<ShapeEx> FromGroup(P.GroupShape xmlGroupTypeShape)
        {
            _currentSdkGroupType = xmlGroupTypeShape;
            var allXmlElements = xmlGroupTypeShape.Elements<OpenXmlCompositeElement>();
            var shapesCollection = new List<ShapeEx>(allXmlElements.Count());
            ParseElements(shapesCollection, allXmlElements);

            return shapesCollection;
        }

        private void ParseElements(List<ShapeEx> shapesCollection, IEnumerable<OpenXmlCompositeElement> allXmlElements)
        {
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
        }

        private void AddOleObjects(List<ShapeEx> shapesCollection, IEnumerable<P.GraphicFrame> xmlOleGraphicFrames)
        {
            foreach (var xmlGraphicFrame in xmlOleGraphicFrames)
            {
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGraphicFrame, _xmlSldPart);
                var ole = new OleObject(xmlGraphicFrame);
                IInnerTransform innerTransform;
                if (_currentSdkGroupType is P.GroupShape groupShape)
                {
                    innerTransform = new NonPlaceholderGroupedTransform(xmlGraphicFrame, groupShape);
                }
                else
                {
                    innerTransform = new NonPlaceholderTransform(xmlGraphicFrame);
                }
                
                var newOleShape = _shapeBuilder.WithOle(innerTransform, spContext, ole);

                shapesCollection.Add(newOleShape);
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
                var picture = new Picture(_xmlSldPart, blipRelateId);

                // Context
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlPic, _xmlSldPart);

                // Location 
                IInnerTransform innerTransform;
                var t2d = xmlPic.ShapeProperties.Transform2D;
                if (t2d != null)
                {
                    if (_currentSdkGroupType is P.GroupShape groupShape)
                    {
                        innerTransform = new NonPlaceholderGroupedTransform(t2d, groupShape);
                    }
                    else
                    {
                        innerTransform = new NonPlaceholderTransform(t2d);
                    }
                }
                else
                {
                    var placeholderLocationData = _phService.TryGet(xmlPic);
                    innerTransform = new PlaceholderTransform(placeholderLocationData);
                }

                var newPicShape = _shapeBuilder.WithPicture(innerTransform, spContext, picture);
                shapesCollection.Add(newPicShape);
            }
        }

        private void AddAutoShapes(List<ShapeEx> shapesCollection, IEnumerable<P.Shape> xmlShapes)
        {
            foreach (var xmlShape in xmlShapes)
            {
                ShapeEx newAutoShape;
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlShape, _xmlSldPart);
                var t2d = xmlShape.ShapeProperties.Transform2D;
                IInnerTransform innerTransform;
                if (t2d != null)
                {
                    if (_currentSdkGroupType is P.GroupShape groupShape)
                    {
                        innerTransform = new NonPlaceholderGroupedTransform(t2d, groupShape);
                    }
                    else
                    {
                        innerTransform = new NonPlaceholderTransform(t2d);
                    }
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

        private void AddTables(List<ShapeEx> shapesCollection, IEnumerable<P.GraphicFrame> xmlTablesGraphicFrames)
        {
            foreach (var xmlGraphicFrame in xmlTablesGraphicFrames)
            {
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGraphicFrame, _xmlSldPart);
                var table = new TableEx(xmlGraphicFrame, spContext);
                IInnerTransform innerTransform;
                if (_currentSdkGroupType is P.GroupShape groupShape)
                {
                    innerTransform = new NonPlaceholderGroupedTransform(xmlGraphicFrame, groupShape);
                }
                else
                {
                    innerTransform = new NonPlaceholderTransform(xmlGraphicFrame);
                }
                var newShape = _shapeBuilder.WithTable(innerTransform, spContext, table);
                shapesCollection.Add(newShape);
            }
        }

        private void AddCharts(List<ShapeEx> shapesCollection, IEnumerable<P.GraphicFrame> xmlTablesGraphicFrames)
        {
            foreach (var xmlGraphicFrame in xmlTablesGraphicFrames)
            {
                var chartEx = new ChartEx(xmlGraphicFrame, _xmlSldPart);
                IInnerTransform innerTransform;
                if (_currentSdkGroupType is P.GroupShape groupShape)
                {
                    innerTransform = new NonPlaceholderGroupedTransform(xmlGraphicFrame, groupShape);
                }
                else
                {
                    innerTransform = new NonPlaceholderTransform(xmlGraphicFrame);
                }
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGraphicFrame, _xmlSldPart);
                var newShape = _shapeBuilder.WithChart(innerTransform, spContext, chartEx);
                shapesCollection.Add(newShape);
            }
        }

        private void AddGroups(List<ShapeEx> shapesCollection, IEnumerable<P.GroupShape> xmlGroupShapes)
        {
            foreach (var xmlGroup in xmlGroupShapes)
            {
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGroup, _xmlSldPart);
                var groupedShapes = FromGroup(xmlGroup);
                var transformGroup = xmlGroup.GroupShapeProperties.TransformGroup;
                var innerTransform = new NonPlaceholderTransform(transformGroup);

                var newShape = _shapeBuilder.WithGroup(innerTransform, spContext, groupedShapes);
                shapesCollection.Add(newShape);
            }
        }

        #endregion Private Methods
    }
}
