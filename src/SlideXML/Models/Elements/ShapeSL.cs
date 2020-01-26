using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Enums;
using SlideXML.Exceptions;
using SlideXML.Extensions;
using SlideXML.Models.Settings;
using SlideXML.Models.TextBody;
using SlideXML.Services;
using SlideXML.Services.Builders;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Models.Elements
{
    /// <summary>
    /// Represents a shape element on a slide.
    /// </summary>
    public class ShapeSL
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private readonly OpenXmlCompositeElement _compositeElement;
        
        private readonly ElementSettings _spSettings;
        private bool? _hidden;
        private int _id;
        private string _name;
        private TextBodySL _textBody;
        private ImageEx _backgroundImage;
        private P.PlaceholderValues? _placeholderType;

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

        public P.PlaceholderValues PlaceholderType
        {
            get
            {
                if (_placeholderType == null)
                {
                    _placeholderType = _compositeElement.GetPlaceholderType();
                }
                if (_placeholderType == null)
                {
                    throw new SlideXMLException("Shape is not a placeholder.");
                }

                return (P.PlaceholderValues)_placeholderType;
            }
        }

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
        /// Returns text body if it exist or null.
        /// </summary>
        /// <remarks>Lazy load.</remarks>
        public TextBodySL TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    TryParseTxtBody();
                }

                return _textBody;
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
        /// Determines whether the shape has text body.
        /// </summary>
        public bool HasTextBody => TextBody != null;

        /// <summary>
        /// Determines whether the shape is placeholder.
        /// </summary>
        public bool IsPlaceholder
        {
            get
            {
                if (_placeholderType == null)
                {
                    _placeholderType = _compositeElement.GetPlaceholderType();
                }

                return _placeholderType != null;
            }
        }


        public ChartSL Chart { get; }

        public TableSL Table { get; }

        public PictureSL Picture { get; }

        public GroupSL Group { get; } //TODO: convert into GroupItems

        public OleObjectSL OleObject { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeSL"/> class.
        /// </summary>
        [SuppressMessage("ReSharper", "SuggestBaseTypeForParameter")]
        private ShapeSL(IBackgroundImageFactory bgImgFactory,
                         OpenXmlCompositeElement compositeElement,
                         SlidePart sldPart,
                         ElementSettings spSettings)
        {
            Check.NotNull(bgImgFactory, nameof(bgImgFactory));
            Check.NotNull(sldPart, nameof(sldPart));
            Check.NotNull(spSettings, nameof(spSettings));
            _bgImgFactory = bgImgFactory;
            _sldPart = sldPart;
            _spSettings = spSettings;
            Type = ShapeType.AutoShape;
            _compositeElement = compositeElement;
        }

        [SuppressMessage("ReSharper", "SuggestBaseTypeForParameter")]
        private ShapeSL(ChartSL chart, OpenXmlCompositeElement ce)
        {
            Chart = chart;
            Type = ShapeType.Chart;
            _compositeElement = ce;
        }

        private ShapeSL(TableSL table, OpenXmlCompositeElement ce)
        {
            Table = table;
            Type = ShapeType.Table;
            _compositeElement = ce;
        }

        private ShapeSL(PictureSL pic, OpenXmlCompositeElement ce)
        {
            Picture = pic;
            Type = ShapeType.Picture;
            _compositeElement = ce;
        }

        private ShapeSL(GroupSL group, OpenXmlCompositeElement ce)
        {
            Group = group;
            Type = ShapeType.Group;
            _compositeElement = ce;
        }

        private ShapeSL(OleObjectSL oleObject, OpenXmlCompositeElement ce)
        {
            OleObject = oleObject;
            Type = ShapeType.OLEObject;
            _compositeElement = ce;
        }

        #endregion Constructors

        #region Private Methods

        private void TryParseTxtBody()
        {
            var pTxtBody = ((P.Shape)_compositeElement).TextBody;

            if (pTxtBody == null)
            {
                return;
            }

            var aTexts = pTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                _textBody = new TextBodySL(_spSettings, pTxtBody);
            }
        }

        private void InitIdHiddenName()
        {
            if (_id == 0) // id == 0: it is mean NonVisualDrawingProperties was not parsed before
            {
                var (id, hidden, name) = _compositeElement.GetNvPrValues();
                _id = id;
                _hidden = hidden;
                _name = name;
            }
        }

        #endregion

        #region Builder

        /// <summary>
        /// Represents slide shape builder.
        /// </summary>
        public class Builder : IShapeNewBuilder
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
            /// Builds shape.
            /// </summary>
            /// <returns></returns>
            public ShapeSL BuildTxtShape(OpenXmlCompositeElement compositeElement, ElementSettings spSettings)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                return new ShapeSL(_bgImgFactor, compositeElement, _sldPart, spSettings);
            }

            public ShapeSL BuildChartShape(P.GraphicFrame xmlGrFrame)
            {
                var chart = new ChartSL(xmlGrFrame, _sldPart);
                var newShape = new ShapeSL(chart, xmlGrFrame);
                var transform = xmlGrFrame.Transform;
                newShape.X = transform.Offset.X.Value;
                newShape.Y = transform.Offset.Y.Value;
                newShape.Width = transform.Extents.Cx.Value;
                newShape.Height = transform.Extents.Cy.Value;

                return newShape;
            }

            public ShapeSL BuildTableShape(P.GraphicFrame xmlGrFrame, ElementSettings elSettings)
            {
                var table = new TableSL(xmlGrFrame, elSettings);
                var newShape = new ShapeSL(table, xmlGrFrame);
                var transform = xmlGrFrame.Transform;
                newShape.X = transform.Offset.X.Value;
                newShape.Y = transform.Offset.Y.Value;
                newShape.Width = transform.Extents.Cx.Value;
                newShape.Height = transform.Extents.Cy.Value;

                return newShape;
            }

            public ShapeSL BuildPictureShape(OpenXmlCompositeElement ce)
            {
                var picture = new PictureSL(_sldPart, ce);
                var newShape = new ShapeSL(picture, ce);
                var transform2D = ce.GetFirstChild<P.ShapeProperties>().Transform2D;
                newShape.X = transform2D.Offset.X.Value;
                newShape.Y = transform2D.Offset.Y.Value;
                newShape.Width = transform2D.Extents.Cx.Value;
                newShape.Height = transform2D.Extents.Cy.Value;

                return newShape;
            }

            public ShapeSL BuildGroupShape(IElementFactory elFactory, OpenXmlCompositeElement ce, IPreSettings preSettings)
            {
                var group = new GroupSL(_groupSpTypeParser, elFactory, ce, preSettings, _sldPart);
                var newShape = new ShapeSL(group, ce);
                var transformGroup = ((P.GroupShape)ce).GroupShapeProperties.TransformGroup;
                newShape.X = transformGroup.Offset.X.Value;
                newShape.Y = transformGroup.Offset.Y.Value;
                newShape.Width = transformGroup.Extents.Cx.Value;
                newShape.Height = transformGroup.Extents.Cy.Value;

                return newShape;
            }

            public ShapeSL BuildOLEObject(OpenXmlCompositeElement ce)
            {
                var oleObject = new OleObjectSL(ce);
                var newShape = new ShapeSL(oleObject, ce);

                return newShape;
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}