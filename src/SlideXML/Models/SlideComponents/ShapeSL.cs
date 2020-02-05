using System;
using System.Diagnostics.CodeAnalysis;
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

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a shape element on a slide.
    /// </summary>
    public class ShapeSL : TaggedComponent
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private readonly OpenXmlCompositeElement _compositeElement;
        private readonly ElementSettings _spSettings;
        private readonly Lazy<ITextFrame> _textFrame = new Lazy<ITextFrame>(new NoTextFrame());
        private readonly Lazy<IChart> _chart = new Lazy<IChart>(new NoChart());
        private bool? _hidden;
        private int _id;
        private string _name;
        private ImageEx _backgroundImage;
        private PlaceholderType? _placeholderType;

        #endregion Fields

        #region Dependencies

        private readonly IBackgroundImageFactory _bgImgFactory;

        #endregion Dependencies

        #region Properties

        public long X { get; set; } //TODO: remove public setter

        public long Y { get; set; }

        public long Width { get; set; }

        public long Height { get; set; }

        public ShapeType Type { get; }

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
        public bool HasChart => Chart is ChartSL;

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
        public TableSL Table { get; }

        /// <summary>
        /// Returns picture.
        /// </summary>
        public PictureSL Picture { get; }

        /// <summary>
        /// Returns group.
        /// </summary>
        public GroupSL Group { get; } //TODO: convert into GroupItems

        /// <summary>
        /// Returns OLE object.
        /// </summary>
        public OleObjectSL OleObject { get; }

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
                    throw new SlideXMLException("Shape is not a placeholder.");
                }

                return (PlaceholderType)_placeholderType;
            }
        }

        #endregion Properties

        #region Constructors

        private ShapeSL(IBackgroundImageFactory bgImgFactory,
                        OpenXmlCompositeElement compositeElement,
                        SlidePart sldPart,
                        ElementSettings spSettings)
        {
            Type = ShapeType.AutoShape;
            _textFrame = new Lazy<ITextFrame>(GetTextFrame);
            _bgImgFactory = bgImgFactory;
            _compositeElement = compositeElement;
            _sldPart = sldPart;
            _spSettings = spSettings;
        }

        private ShapeSL(IChart chart, OpenXmlCompositeElement ce)
        {
            Type = ShapeType.Chart;
            _chart = new Lazy<IChart>(chart);
            _compositeElement = ce;
        }

        private ShapeSL(TableSL table, OpenXmlCompositeElement ce)
        {
            Type = ShapeType.Table;
            Table = table;
            _compositeElement = ce;
        }

        private ShapeSL(PictureSL pic, OpenXmlCompositeElement ce)
        {
            Type = ShapeType.Picture;
            Picture = pic;
            _compositeElement = ce;
        }

        private ShapeSL(GroupSL group, OpenXmlCompositeElement ce)
        {
            Type = ShapeType.Group;
            Group = group;
            _compositeElement = ce;
        }

        private ShapeSL(OleObjectSL oleObject, OpenXmlCompositeElement ce)
        {
            Type = ShapeType.OLEObject;
            OleObject = oleObject;
            _compositeElement = ce;
        }

        #endregion Constructors

        #region Private Methods

        private ITextFrame GetTextFrame()
        {
            if (Type != ShapeType.AutoShape)
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
            private readonly IGroupShapeTypeParser _groupSpTypeParser;
            private readonly SlidePart _sldPart;

            #endregion Dependencies

            #region Constructors

            /// <summary>
            /// Initialize a new builder.
            /// </summary>
            /// <param name="bgImgFactor"></param>
            /// <param name="groupSpTypeParser"></param>
            public Builder(IBackgroundImageFactory bgImgFactor, IGroupShapeTypeParser groupSpTypeParser, SlidePart sldPart)
            {
                Check.NotNull(bgImgFactor, nameof(bgImgFactor));
                _bgImgFactor = bgImgFactor;
                _groupSpTypeParser = groupSpTypeParser;
                _sldPart = sldPart;
            }

            #endregion Constructors

            #region Public Methods

            /// <summary>
            /// Builds a AutoShape.
            /// </summary>
            /// <returns></returns>
            public ShapeSL BuildAutoShape(OpenXmlCompositeElement compositeElement, ElementSettings spSettings)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                Check.NotNull(spSettings, nameof(spSettings));

                return new ShapeSL(_bgImgFactor, compositeElement, _sldPart, spSettings);
            }

            /// <summary>
            /// Builds a chart.
            /// </summary>
            /// <param name="xmlGrFrame"></param>
            /// <returns></returns>
            public ShapeSL BuildChart(P.GraphicFrame xmlGrFrame)
            {
                Check.NotNull(xmlGrFrame, nameof(xmlGrFrame));
                var newShape = new ShapeSL(new ChartSL(xmlGrFrame, _sldPart), xmlGrFrame);
                SetTransform(newShape, xmlGrFrame);

                return newShape;
            }

            /// <summary>
            /// Builds a table.
            /// </summary>
            /// <param name="xmlGrFrame"></param>
            /// <param name="elSettings"></param>
            /// <returns></returns>
            public ShapeSL BuildTable(P.GraphicFrame xmlGrFrame, ElementSettings elSettings)
            {
                Check.NotNull(xmlGrFrame, nameof(xmlGrFrame));
                Check.NotNull(elSettings, nameof(elSettings));
                var newShape = new ShapeSL(new TableSL(xmlGrFrame, elSettings), xmlGrFrame);
                SetTransform(newShape, xmlGrFrame);

                return newShape;
            }

            /// <summary>
            /// Builds a picture.
            /// </summary>
            /// <returns></returns>
            public ShapeSL BuildPicture(OpenXmlCompositeElement compositeElement)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                var newShape = new ShapeSL(new PictureSL(_sldPart, compositeElement), compositeElement);
                SetTransform2D(newShape, compositeElement);

                return newShape;
            }

            /// <summary>
            /// Builds a OLE object.
            /// </summary>
            /// <returns></returns>
            public ShapeSL BuildOLEObject(OpenXmlCompositeElement compositeElement)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                var newShape = new ShapeSL(new OleObjectSL(compositeElement), compositeElement);
                SetTransform2D(newShape, compositeElement);

                return newShape;
            }

            /// <summary>
            /// Builds a group.
            /// </summary>
            /// <returns></returns>
            public ShapeSL BuildGroup(IElementFactory elFactory, OpenXmlCompositeElement compositeElement, IPreSettings preSettings)
            {
                Check.NotNull(elFactory, nameof(elFactory));
                Check.NotNull(compositeElement, nameof(compositeElement));
                var group = new GroupSL(_groupSpTypeParser, elFactory, compositeElement, preSettings, _sldPart);
                var newShape = new ShapeSL(group, compositeElement);
               
                var transformGroup = ((P.GroupShape)compositeElement).GroupShapeProperties.TransformGroup;
                newShape.X = transformGroup.Offset.X.Value;
                newShape.Y = transformGroup.Offset.Y.Value;
                newShape.Width = transformGroup.Extents.Cx.Value;
                newShape.Height = transformGroup.Extents.Cy.Value;

                return newShape;
            }

            #endregion Public Methods

            private static void SetTransform(ShapeSL newShape, OpenXmlElement ce)
            {
                var transform = ce.Descendants<P.Transform>().Single();
                newShape.X = transform.Offset.X.Value;
                newShape.Y = transform.Offset.Y.Value;
                newShape.Width = transform.Extents.Cx.Value;
                newShape.Height = transform.Extents.Cy.Value;
            }

            private static void SetTransform2D(ShapeSL newShape, OpenXmlElement ce)
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