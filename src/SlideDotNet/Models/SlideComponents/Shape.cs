using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Exceptions;
using SlideDotNet.Extensions;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents.Chart;
using SlideDotNet.Models.TextBody;
using SlideDotNet.Services;
using SlideDotNet.Services.Builders;
using SlideDotNet.Services.Placeholders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using Table = SlideXML.Models.SlideComponents.Table;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models.SlideComponents
{
    /// <summary>
    /// Represents a slide element.
    /// </summary>
    public class Shape
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private readonly OpenXmlCompositeElement _compositeElement;
        private readonly ElementSettings _elSettings;
        private readonly Lazy<ITextFrame> _textFrame = new Lazy<ITextFrame>(new NoTextFrame());
        private readonly Lazy<IChart> _chart = new Lazy<IChart>(new NoChart());
        private readonly IBackgroundImageFactory _bgImgFactory;
        private bool? _hidden;
        private int _id;
        private string _name;
        private ImageEx _backgroundImage;
        private PlaceholderType? _placeholderType;
        private readonly Picture _picture;

        #endregion Fields

        #region Properties

        public long X { get; set; } //TODO: remove public setter

        public long Y { get; set; }

        public long Width { get; set; }

        public long Height { get; set; }

        public ElementType Type { get; }

        /// <summary>
        /// Returns an element identifier.
        /// </summary>
        public int Id
        {
            get
            {
                InitIdHiddenName();
                return _id;
            }
        }

        /// <summary>
        /// Gets an element name.
        /// </summary>
        public string Name
        {
            get
            {
                InitIdHiddenName();
                return _name;
            }
        }

        /// <summary>
        /// Returns background image of AutoShape. Returns null if AutoShape has not background image.
        /// </summary>
        public ImageEx BackgroundImage
        {
            get
            { //TODO: throw exception if element is not a AutoShape
                return _backgroundImage ??= _bgImgFactory.CreateBackgroundShape(_sldPart, (P.Shape) _compositeElement);
            }
        }

        /// <summary>
        /// Determines whether the shape is hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                InitIdHiddenName();
                return (bool)_hidden;
            }
        }

        /// <summary>
        /// Determines whether the shape has text frame.
        /// </summary>
        public bool HasTextFrame => TextFrame is TextFrame;

        /// <summary>
        /// Determines whether the shape has chart.
        /// </summary>
        public bool HasChart => Chart is Chart.Chart;

        /// <summary>
        /// Determines whether the slide element has picture.
        /// </summary>
        public bool HasPicture => _picture != null;

        /// <summary>
        /// Returns text frame.
        /// </summary>
        /// <remarks>Lazy load.</remarks>
        public ITextFrame TextFrame => _textFrame.Value;

        /// <summary>
        /// Returns chart.
        /// </summary>
        public IChart Chart => _chart.Value;

        /// <summary>
        /// Returns table.
        /// </summary>
        public Table Table { get; }

        /// <summary>
        /// Returns picture.
        /// </summary>
        public Picture Picture => _picture ?? throw new SlideXmlException(ExceptionMessages.NoPicture);

        public IList<Shape> GroupedElements { get; }

        /// <summary>
        /// Returns OLE object.
        /// </summary>
        public OleObject OleObject { get; }

        /// <summary>
        /// Determines shape is placeholder.
        /// </summary>
        public bool IsPlaceholder => _compositeElement.IsPlaceholder(); //TODO: make lazy

        public PlaceholderType PlaceholderType
        {
            get
            {
                if (_placeholderType == null && _compositeElement.IsPlaceholder())
                {
                    _placeholderType = PlaceholderService.XmlPlaceholderFrom(_compositeElement).PlaceholderType;
                }
                if (_placeholderType == null)
                {
                    throw new SlideXmlException(ExceptionMessages.NotPlaceholder);
                }

                return (PlaceholderType)_placeholderType;
            }
        }

        #endregion Properties

        #region Constructors

        private Shape(IBackgroundImageFactory bgImgFactory,
                             OpenXmlCompositeElement compositeElement,
                             SlidePart sldPart, 
                             ElementSettings elSettings)
        {

            _bgImgFactory = bgImgFactory;
            _compositeElement = compositeElement;
            _sldPart = sldPart;
            _elSettings = elSettings;
            _elSettings.SlideElement = this;
            Type = ElementType.AutoShape;
            _textFrame = new Lazy<ITextFrame>(GetTextFrame);
        }

        private Shape(IChart chart, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Chart;
            _chart = new Lazy<IChart>(chart);
            _compositeElement = ce;
        }

        private Shape(Table table, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Table;
            Table = table;
            _compositeElement = ce;
        }

        private Shape(Picture pic, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Picture;
            _picture = pic;
            _compositeElement = ce;
        }

        private Shape(Group group, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Group;
            GroupedElements = new List<Shape>(group.Shapes);
            _compositeElement = ce;
        }

        private Shape(OleObject oleObject, OpenXmlCompositeElement ce)
        {
            Type = ElementType.OLEObject;
            OleObject = oleObject;
            _compositeElement = ce;
        }

        #endregion Constructors

        #region Private Methods

        private ITextFrame GetTextFrame()
        {
            if (Type != ElementType.AutoShape)
            {
                return new NoTextFrame();
            }

            var pTxtBody = _compositeElement.Descendants<P.TextBody>().SingleOrDefault();
            if (pTxtBody == null)
            {
                return new NoTextFrame();
            }

            var aTexts = pTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                return new TextFrame(_elSettings, pTxtBody);
            }

            return new NoTextFrame();
        }

        private void InitIdHiddenName()
        {
            if (_id != 0)
            {
                return;
            }
            var (id, hidden, name) = _compositeElement.GetNvPrValues();
            _id = id;
            _hidden = hidden;
            _name = name;
        }

        #endregion

        #region Builder

        /// <inheritdoc cref="IShapeBuilder"/>
        public class Builder : IShapeBuilder
        {
            #region Dependencies

            private readonly IBackgroundImageFactory _bgImgFactor;
            private readonly IXmlGroupShapeTypeParser _groupSpTypeParser;
            private readonly SlidePart _sldPart;
            private readonly PlaceholderService _phService;

            #endregion Dependencies

            #region Constructors

            /// <summary>
            /// Initialize a new builder.
            /// </summary>
            public Builder(IBackgroundImageFactory bgImgFactor, IXmlGroupShapeTypeParser groupSpTypeParser, SlidePart sldPart)
            {
                _bgImgFactor = bgImgFactor ?? throw new ArgumentNullException(nameof(bgImgFactor));
                _groupSpTypeParser = groupSpTypeParser ?? throw new ArgumentNullException(nameof(groupSpTypeParser));
                _sldPart = sldPart ?? throw new ArgumentNullException(nameof(sldPart));
                _phService = new PlaceholderService(_sldPart.SlideLayoutPart);
            }

            #endregion Constructors

            #region Public Methods

            /// <summary>
            /// Builds a AutoShape.
            /// </summary>
            /// <returns></returns>
            public Shape BuildAutoShape(OpenXmlCompositeElement xmlElement, ElementSettings elSettings)
            {
                Check.NotNull(xmlElement, nameof(xmlElement));
                Check.NotNull(elSettings, nameof(elSettings));

                Shape shape;
                var t2d = ((P.Shape)xmlElement).ShapeProperties.Transform2D;
                if (t2d != null)
                {
                    if (xmlElement.IsPlaceholder())
                    {
                        elSettings.Placeholder = _phService.TryGet(xmlElement);
                    }
                    // creates AutoShape
                    shape = new Shape(_bgImgFactor, xmlElement, _sldPart, elSettings);
                    elSettings.SlideElement = shape;
                    WithOwnTransform2d(shape, t2d);
                }
                else // is placeholder obviously
                {
                    var placeholder = _phService.TryGet(xmlElement);
                    elSettings.Placeholder = placeholder;

                    shape = new Shape(_bgImgFactor, xmlElement, _sldPart, elSettings);
                    elSettings.SlideElement = shape;
                    shape.X = placeholder.X;
                    shape.Y = placeholder.Y;
                    shape.Width = placeholder.Width;
                    shape.Height = placeholder.Height;
                }

                return shape;
            }

            /// <summary>
            /// Builds a picture.
            /// </summary>
            /// <returns></returns>
            public Shape BuildPicture(OpenXmlCompositeElement xmlElement, ElementSettings elSettings)
            {
                Check.NotNull(xmlElement, nameof(xmlElement));
                Check.NotNull(elSettings, nameof(elSettings));

                Shape shape;
                var t2d = ((P.Picture)xmlElement).ShapeProperties.Transform2D;
                if (t2d != null)
                {
                    if (xmlElement.IsPlaceholder())
                    {
                        elSettings.Placeholder = _phService.TryGet(xmlElement);
                    }
                    // creates Picture
                    shape = new Shape(new Picture(_sldPart, xmlElement), xmlElement);

                    elSettings.SlideElement = shape;
                    WithOwnTransform2d(shape, t2d);
                }
                else // is placeholder obviously
                {
                    var placeholder = _phService.TryGet(xmlElement);
                    elSettings.Placeholder = placeholder;

                    shape = new Shape(new Picture(_sldPart, xmlElement), xmlElement);
                    elSettings.SlideElement = shape;
                    shape.X = placeholder.X;
                    shape.Y = placeholder.Y;
                    shape.Width = placeholder.Width;
                    shape.Height = placeholder.Height;
                }

                return shape;

                //var newShape = new Shape(new Picture(_sldPart, compositeElement), compositeElement);
                //SetTransform2D(newShape, compositeElement);
                //return shape;
            }

            /// <summary>
            /// Builds a chart.
            /// </summary>
            /// <param name="xmlGrFrame"></param>
            /// <returns></returns>
            public Shape BuildChart(P.GraphicFrame xmlGrFrame)
            {
                Check.NotNull(xmlGrFrame, nameof(xmlGrFrame));
                var newShape = new Shape(new Chart.Chart(xmlGrFrame, _sldPart), xmlGrFrame);
                SetTransform(newShape, xmlGrFrame);

                return newShape;
            }

            /// <summary>
            /// Builds a table.
            /// </summary>
            /// <param name="xmlGrFrame"></param>
            /// <param name="elSettings"></param>
            /// <returns></returns>
            public Shape BuildTable(P.GraphicFrame xmlGrFrame, ElementSettings elSettings)
            {
                Check.NotNull(xmlGrFrame, nameof(xmlGrFrame));
                Check.NotNull(elSettings, nameof(elSettings));
                var newShape = new Shape(new Table(xmlGrFrame, elSettings), xmlGrFrame);
                SetTransform(newShape, xmlGrFrame);

                return newShape;
            }

            /// <summary>
            /// Builds a OLE object.
            /// </summary>
            /// <returns></returns>
            public Shape BuildOleObject(OpenXmlCompositeElement xmlElement)
            {
                Check.NotNull(xmlElement, nameof(xmlElement));
                var newShape = new Shape(new OleObject(xmlElement), xmlElement);
                SetTransform2D(newShape, xmlElement);

                return newShape;
            }

            /// <summary>
            /// Builds a group.
            /// </summary>
            /// <returns></returns>
            public Shape BuildGroup(IElementFactory elFactory, OpenXmlCompositeElement xmlElement, IParents parents)
            {
                Check.NotNull(elFactory, nameof(elFactory));
                Check.NotNull(xmlElement, nameof(xmlElement));
                var group = new Group(_groupSpTypeParser, elFactory, xmlElement, parents);
                var newShape = new Shape(group, xmlElement);
               
                var transformGroup = ((P.GroupShape)xmlElement).GroupShapeProperties.TransformGroup;
                newShape.X = transformGroup.Offset.X.Value;
                newShape.Y = transformGroup.Offset.Y.Value;
                newShape.Width = transformGroup.Extents.Cx.Value;
                newShape.Height = transformGroup.Extents.Cy.Value;

                return newShape;
            }

            #endregion Public Methods

            private static void SetTransform(Shape shape, OpenXmlElement xmlElement)
            {
                var transform = xmlElement.Descendants<P.Transform>().Single();
                shape.X = transform.Offset.X.Value;
                shape.Y = transform.Offset.Y.Value;
                shape.Width = transform.Extents.Cx.Value;
                shape.Height = transform.Extents.Cy.Value;
            }

            private static void SetTransform2D(Shape shape, OpenXmlElement xmlElement)
            {
                var transform2D = xmlElement.Descendants<A.Transform2D>().Single();
                WithOwnTransform2d(shape, transform2D);
            }

            private static void WithOwnTransform2d(Shape shape, A.Transform2D t2D)
            {
                shape.X = t2D.Offset.X.Value;
                shape.Y = t2D.Offset.Y.Value;
                shape.Width = t2D.Extents.Cx.Value;
                shape.Height = t2D.Extents.Cy.Value;
            }
        }

        #endregion Builder
    }
}