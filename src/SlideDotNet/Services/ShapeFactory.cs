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

        private readonly IShapeBuilder _shapeBuilder = new ShapeEx.Builder(); // make DI
        private readonly IPreSettings _preSettings;
        private readonly IPlaceholderService _phService;
        private readonly SlidePlaceholderFontService _slideFontService;
        private readonly SlidePart _xmlSldPart;

        #endregion Fields

        public ShapeFactory(SlidePart xmlSldPart, IPreSettings preSettings)
        {
            _xmlSldPart = xmlSldPart ?? throw new ArgumentNullException(nameof(xmlSldPart));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _slideFontService = new SlidePlaceholderFontService(xmlSldPart);
            _phService = new PlaceholderService(xmlSldPart.SlideLayoutPart);
        }

        #region Public Methods

        public IList<ShapeEx> CreateShapesCollection(P.ShapeTree xmlGroupTypeShape)
        {
            var allXmlElements = xmlGroupTypeShape.Elements<OpenXmlCompositeElement>();
            var shapesCollection = new List<ShapeEx>();
            ParseTree(shapesCollection, allXmlElements);

            var xmlGroups = allXmlElements.OfType<P.GroupShape>();
            AddGroups(shapesCollection, xmlGroups);

            return shapesCollection;
        }

        #endregion Public Methods

        #region Private Methods

        private IList<ShapeEx> CreateShapesCollection(P.GroupShape xmlGroupTypeShape)
        {
            var allXmlElements = xmlGroupTypeShape.Elements<OpenXmlCompositeElement>();
            var shapesCollection = new List<ShapeEx>();
            ParseTree(shapesCollection, allXmlElements);

            return shapesCollection;
        }

        private void ParseTree(List<ShapeEx> shapesCollection, IEnumerable<OpenXmlCompositeElement> allXmlElements)
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
                var transform = xmlGraphicFrame.Transform;
                var location = LocationFromTransformX(transform);
                var newOleShape = _shapeBuilder.WithOle(location, spContext, ole);

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
                Location location;
                var t2d = xmlPic.ShapeProperties.Transform2D;
                if (t2d != null)
                {
                    location = LocationFromTransformX(t2d);
                }
                else
                {
                    var phData = _phService.TryGet(xmlPic);
                    location = LocationFromPh(phData);
                }

                var newPicShape = _shapeBuilder.WithPicture(location, spContext, picture);
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
                if (t2d != null)
                {
                    if (xmlShape.IsPlaceholder())
                    {
                        spContext.PlaceholderLocationData = _phService.TryGet(xmlShape);
                    }

                    var location = LocationFromTransformX(t2d);
                    newAutoShape = _shapeBuilder.WithAutoShape(location, spContext);
                }
                else // is placeholder obviously
                {
                    var placeholder = _phService.TryGet(xmlShape);
                    spContext.PlaceholderLocationData = placeholder;

                    var location = new Location
                    {
                        X = placeholder.X,
                        Y = placeholder.Y,
                        Width = placeholder.Width,
                        Height = placeholder.Height
                    };
                    newAutoShape = _shapeBuilder.WithAutoShape(location, spContext);
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
                var transform = xmlGraphicFrame.Transform;
                var location = LocationFromTransformX(transform);
                var newShape = _shapeBuilder.WithTable(location, spContext, table);
                shapesCollection.Add(newShape);
            }
        }

        private void AddCharts(List<ShapeEx> shapesCollection, IEnumerable<P.GraphicFrame> xmlTablesGraphicFrames)
        {
            foreach (var xmlGraphicFrame in xmlTablesGraphicFrames)
            {
                var chart = new ChartEx(xmlGraphicFrame, _xmlSldPart);
                var transform = xmlGraphicFrame.Transform;
                var location = LocationFromTransformX(transform);
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGraphicFrame, _xmlSldPart);
                var newShape = _shapeBuilder.WithChart(location, spContext, chart);
                shapesCollection.Add(newShape);
            }
        }

        private void AddGroups(List<ShapeEx> shapeCollection, IEnumerable<P.GroupShape> xmlGroupShapes)
        {
            foreach (var xmlGroup in xmlGroupShapes)
            {
                var spContext = new ShapeContext(_preSettings, _slideFontService, xmlGroup, _xmlSldPart);
                var groupedShapes = CreateShapesCollection(xmlGroup);
                var transformGroup = xmlGroup.GroupShapeProperties.TransformGroup;
                foreach (var item in groupedShapes)
                {
                    item.X = item.X - transformGroup.ChildOffset.X + transformGroup.Offset.X;
                    item.Y = item.Y - transformGroup.ChildOffset.Y + transformGroup.Offset.Y;
                }
                var location = new Location
                {
                    X = transformGroup.Offset.X.Value,
                    Y = transformGroup.Offset.Y.Value,
                    Width = transformGroup.Extents.Cx.Value,
                    Height = transformGroup.Extents.Cy.Value
                };
                var newShape = _shapeBuilder.WithGroup(location, spContext, groupedShapes);
                shapeCollection.Add(newShape);
            }
        }

        private static Location LocationFromTransformX(OpenXmlCompositeElement xmlElement)
        {
            var offset = xmlElement.GetFirstChild<A.Offset>();
            var extents = xmlElement.GetFirstChild<A.Extents>();

            return new Location
            {
                X = offset.X.Value,
                Y = offset.Y.Value,
                Width = extents.Cx.Value,
                Height = extents.Cy.Value
            };
        }

        private static Location LocationFromPh(PlaceholderLocationData phLocationData)
        {
            return new Location
            {
                X = phLocationData.X,
                Y = phLocationData.Y,
                Width = phLocationData.Width,
                Height = phLocationData.Height
            };
        }

        #endregion Private Methods
    }
}
