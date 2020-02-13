using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Enums;
using SlideXML.Exceptions;
using SlideXML.Extensions;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents.Chart;
using SlideXML.Models.TextBody;
using SlideXML.Services;
using SlideXML.Services.Builders;
using SlideXML.Services.Placeholders;
using SlideXML.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a slide element.
    /// </summary>
    public class SlideElement
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private readonly OpenXmlCompositeElement _compositeElement;
        private readonly ElementSettings _spSettings;
        private readonly Lazy<ITextFrame> _textFrame = new Lazy<ITextFrame>(new NoTextFrame());
        private readonly Lazy<IChart> _chart = new Lazy<IChart>(new NoChart());
        private readonly IBackgroundImageFactory _bgImgFactory;
        private bool? _hidden;
        private int _id;
        private string _name;
        private ImageEx _backgroundImage;
        private PlaceholderType? _placeholderType;

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
        /// Gets background image.
        /// </summary>
        /// <returns><see cref="ImageEx"/> instance or null if shape has not background image.</returns>
        public ImageEx BackgroundImage
        {
            get
            {
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
        public Picture Picture { get; }

        public IList<SlideElement> GroupedElements { get; }

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
                    _placeholderType = PlaceholderService.GetPlaceholderXML(_compositeElement).PlaceholderType;
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

        private SlideElement(IBackgroundImageFactory bgImgFactory,
                             OpenXmlCompositeElement compositeElement,
                             SlidePart sldPart, 
                             ElementSettings spSettings)
        {
            
            _bgImgFactory = bgImgFactory ?? throw new ArgumentNullException(nameof(bgImgFactory));
            _compositeElement = compositeElement ?? throw new ArgumentNullException(nameof(compositeElement));
            _sldPart = sldPart ?? throw new ArgumentNullException(nameof(sldPart));
            _spSettings = spSettings;
            _spSettings.Shape = this;
            Type = ElementType.AutoShape;
            _textFrame = new Lazy<ITextFrame>(GetTextFrame);
        }

        private SlideElement(IChart chart, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Chart;
            _chart = new Lazy<IChart>(chart);
            _compositeElement = ce;
        }

        private SlideElement(Table table, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Table;
            Table = table;
            _compositeElement = ce;
        }

        private SlideElement(Picture pic, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Picture;
            Picture = pic;
            _compositeElement = ce;
        }

        private SlideElement(Group group, OpenXmlCompositeElement ce)
        {
            Type = ElementType.Group;
            GroupedElements = new List<SlideElement>(group.Shapes);
            _compositeElement = ce;
        }

        private SlideElement(OleObject oleObject, OpenXmlCompositeElement ce)
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
                return new TextFrame(_spSettings, pTxtBody);
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

        /// <summary>
        /// Represents slide shape builder.
        /// </summary>
        public class Builder : IShapeBuilder
        {
            #region Dependencies

            private readonly IBackgroundImageFactory _bgImgFactor;
            private readonly IXmlGroupShapeTypeParser _groupSpTypeParser;
            private readonly SlidePart _sldPart;

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
            }

            #endregion Constructors

            #region Public Methods

            /// <summary>
            /// Builds a AutoShape.
            /// </summary>
            /// <returns></returns>
            public SlideElement BuildAutoShape(OpenXmlCompositeElement compositeElement, ElementSettings spSettings)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                Check.NotNull(spSettings, nameof(spSettings));

                return new SlideElement(_bgImgFactor, compositeElement, _sldPart, spSettings);
            }

            /// <summary>
            /// Builds a chart.
            /// </summary>
            /// <param name="xmlGrFrame"></param>
            /// <returns></returns>
            public SlideElement BuildChart(P.GraphicFrame xmlGrFrame)
            {
                Check.NotNull(xmlGrFrame, nameof(xmlGrFrame));
                var newShape = new SlideElement(new Chart.Chart(xmlGrFrame, _sldPart), xmlGrFrame);
                SetTransform(newShape, xmlGrFrame);

                return newShape;
            }

            /// <summary>
            /// Builds a table.
            /// </summary>
            /// <param name="xmlGrFrame"></param>
            /// <param name="elSettings"></param>
            /// <returns></returns>
            public SlideElement BuildTable(P.GraphicFrame xmlGrFrame, ElementSettings elSettings)
            {
                Check.NotNull(xmlGrFrame, nameof(xmlGrFrame));
                Check.NotNull(elSettings, nameof(elSettings));
                var newShape = new SlideElement(new Table(xmlGrFrame, elSettings), xmlGrFrame);
                SetTransform(newShape, xmlGrFrame);

                return newShape;
            }

            /// <summary>
            /// Builds a picture.
            /// </summary>
            /// <returns></returns>
            public SlideElement BuildPicture(OpenXmlCompositeElement compositeElement)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                var newShape = new SlideElement(new Picture(_sldPart, compositeElement), compositeElement);
                SetTransform2D(newShape, compositeElement);

                return newShape;
            }

            /// <summary>
            /// Builds a OLE object.
            /// </summary>
            /// <returns></returns>
            public SlideElement BuildOLEObject(OpenXmlCompositeElement compositeElement)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                var newShape = new SlideElement(new OleObject(compositeElement), compositeElement);
                SetTransform2D(newShape, compositeElement);

                return newShape;
            }

            /// <summary>
            /// Builds a group.
            /// </summary>
            /// <returns></returns>
            public SlideElement BuildGroup(IElementFactory elFactory, OpenXmlCompositeElement compositeElement, IPreSettings preSettings)
            {
                Check.NotNull(elFactory, nameof(elFactory));
                Check.NotNull(compositeElement, nameof(compositeElement));
                var group = new Group(_groupSpTypeParser, elFactory, compositeElement, preSettings, _sldPart);
                var newShape = new SlideElement(group, compositeElement);
               
                var transformGroup = ((P.GroupShape)compositeElement).GroupShapeProperties.TransformGroup;
                newShape.X = transformGroup.Offset.X.Value;
                newShape.Y = transformGroup.Offset.Y.Value;
                newShape.Width = transformGroup.Extents.Cx.Value;
                newShape.Height = transformGroup.Extents.Cy.Value;

                return newShape;
            }

            #endregion Public Methods

            private static void SetTransform(SlideElement newShape, OpenXmlElement ce)
            {
                var transform = ce.Descendants<P.Transform>().Single();
                newShape.X = transform.Offset.X.Value;
                newShape.Y = transform.Offset.Y.Value;
                newShape.Width = transform.Extents.Cx.Value;
                newShape.Height = transform.Extents.Cy.Value;
            }

            private static void SetTransform2D(SlideElement newShape, OpenXmlElement ce)
            {
                var transform2D = ce.Descendants<A.Transform2D>().Single();
                newShape.X = transform2D.Offset.X.Value;
                newShape.Y = transform2D.Offset.Y.Value;
                newShape.Width = transform2D.Extents.Cx.Value;
                newShape.Height = transform2D.Extents.Cy.Value;
            }
        }

        #endregion Builder
    }
}