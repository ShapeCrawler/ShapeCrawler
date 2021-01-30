using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Enums;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories.Builders;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.Styles;
using ShapeCrawler.Models.Transforms;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using ShapeCrawler.Tables;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable CheckNamespace

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    /// Represents a shape on a slide.
    /// </summary>
    public class 
        ShapeSc
    {
        #region Fields

        internal ShapeContext Context;
        private readonly Lazy<TextBoxSc> _text;
        private readonly Lazy<ShapeFill> _shapeFill;
        private readonly IImageExFactory _imageFactory = new ImageExFactory(); //TODO: do not initiate for non-AutoShape types
        private bool? _hidden;
        private int _id;
        private string _name;
        private PictureSc _picture;
        private OLEObject _ole;
        private TableSc _table;
        private ChartSc _chart;
        private readonly ILocation _innerTransform;

        #endregion Fields

        internal OpenXmlCompositeElement ShapeTreeSource { get; set; }

        #region Public Properties

        /// <summary>
        /// Returns the x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X
        {
            get => _innerTransform.X;
            set => _innerTransform.SetX(value);
        }

        /// <summary>
        /// Returns the y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y
        {
            get => _innerTransform.Y;
            set => _innerTransform.SetY(value);
        }

        /// <summary>
        /// Returns the width of the shape.
        /// </summary>
        public long Width
        {
            get => _innerTransform.Width;
            set => _innerTransform.SetWidth(value);
        }

        /// <summary>
        /// Returns the height of the shape.
        /// </summary>
        public long Height
        {
            get => _innerTransform.Height;
            set => _innerTransform.SetHeight(value);
        }

        /// <summary>
        /// Returns shape main content type.
        /// </summary>
        public ShapeContentType ContentType { get; }

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
        public bool HasTextBox => TextBox is TextBoxSc;

        /// <summary>
        /// Determines whether the shape has chart content.
        /// </summary>
        public bool HasChart => Chart is ChartSc;

        /// <summary>
        /// Determines whether the slide element has picture content.
        /// </summary>
        public bool HasPicture => _picture != null;

        /// <summary>
        /// Returns text frame.
        /// </summary>
        /// <remarks>Lazy load.</remarks>
        public TextBoxSc TextBox => _text.Value;

        /// <summary>
        /// Returns chart. Returns <c>NULL</c> when the shape content type is not <see cref="ShapeContentType.Chart"/>.
        /// </summary>
        public ChartSc Chart => _chart;

        /// <summary>
        /// Returns table. Returns <c>NULL</c> when the shape content type is not <see cref="ShapeContentType.Table"/>.
        /// </summary>
        public TableSc Table => _table;

        /// <summary>
        /// Returns picture. Returns <c>NULL</c> when the shape content type is not a <see cref="ShapeContentType.Picture"/>.
        /// </summary>
        public PictureSc Picture => _picture;

        /// <summary>
        /// Returns grouped shapes. Returns <c>NULL</c> when the shape content type is not <see cref="ShapeContentType.Group"/>.
        /// </summary>
        public IList<ShapeSc> GroupedShapes { get; private set; }

        /// <summary>
        /// Returns OLE object content. Returns <c>NULL</c> when the shape content type is not <see cref="ShapeContentType.OLEObject"/>.
        /// </summary>
        public OLEObject OleObject => _ole;

        /// <summary>
        /// Determines whether the shape is placeholder.
        /// </summary>
        public bool IsPlaceholder => Placeholder != null;

        public Placeholder Placeholder
        {
            get
            {
                if (Context.CompositeElement.IsPlaceholder())
                {
                    return new Placeholder();
                }

                return null;
            }
        }

        /// <summary>
        /// Returns placeholder type. Returns null if shape is not a placeholder.
        /// </summary>
        public PlaceholderType? PlaceholderType
        {
            get
            {
                if (!IsPlaceholder)
                {
                    return null;
                }
                
                return Context.PlaceholderService.GetPlaceholderType(Context.CompositeElement);
            }
        }

        /// <summary>
        /// Returns the Shape Fill. Returns <c>NULL</c> if shape is not filled.
        /// </summary>
        public ShapeFill Fill => _shapeFill.Value;

        /// <summary>
        /// Determines whether the shape is grouped.
        /// </summary>
        public bool IsGrouped => _innerTransform is NonPlaceholderGroupedTransform;

        public GeometryType GeometryType { get; }

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        /// <summary>
        /// Gets parent slide.
        /// </summary>
        public SlideSc Slide { get; internal set; }

        #endregion Properties

        #region Constructors

        private ShapeSc(
            ILocation innerTransform,
            ShapeContext spContext,
            ShapeContentType contentType,
            GeometryType geometryType,
            OpenXmlCompositeElement shapeTreeSource) : this(innerTransform, spContext, contentType, shapeTreeSource)
        {
            GeometryType = geometryType;
        }

        private ShapeSc (ILocation innerTransform, ShapeContext spContext, ShapeContentType contentType, OpenXmlCompositeElement shapeTreeSource)
        {
            _innerTransform = innerTransform;
            Context = spContext;
            ContentType = contentType;
            _text = new Lazy<TextBoxSc>(GetTextBox);
            _shapeFill = new Lazy<ShapeFill>(TryGetFill);
            ShapeTreeSource = shapeTreeSource;
        }

        #endregion Constructors

        #region Private Methods

        private void SetCustomData(string value)
        {
            var customDataElement = $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            Context.CompositeElement.InnerXml += customDataElement;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(Context.CompositeElement.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        private TextBoxSc GetTextBox()
        {
            if (ContentType != ShapeContentType.AutoShape)
            {
                return null;
            }

            P.TextBody pTextBody = Context.CompositeElement.Descendants<P.TextBody>().SingleOrDefault();
            if (pTextBody == null)
            {
                return null;
            }

            var aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element with text must be exist
            {
                return new TextBoxSc(this, pTextBody);
            }

            return null;
        }

        private ShapeFill TryGetFill()
        {
            if (ContentType != ShapeContentType.AutoShape)
            {
                return null;
            }
            var image = _imageFactory.TryFromSdkShape(Context.SlidePart, (OpenXmlCompositeElement)Context.CompositeElement); //TODO: delete casting
            if (image != null)
            {
                return new ShapeFill(image);
            }

            var xmlShape = (P.Shape) Context.CompositeElement;
            var rgbColorModelHex = xmlShape.ShapeProperties.GetFirstChild<A.SolidFill>()?.RgbColorModelHex;
            if (rgbColorModelHex != null)
            {
                return ShapeFill.FromXmlSolidFill(rgbColorModelHex);
            }

            return null;
        }

        private void InitIdHiddenName()
        {
            if (_id != 0)
            {
                return;
            }
            var (id, hidden, name) = Context.CompositeElement.GetNvPrValues();
            _id = id;
            _hidden = hidden;
            _name = name;
        }

        #endregion

        #region Builder

        /// <inheritdoc cref="IShapeBuilder"/>
        public class Builder : IShapeBuilder
        {
            #region Public Methods

            public ShapeSc WithOle(ILocation innerTransform, ShapeContext spContext, OLEObject ole, OpenXmlCompositeElement shapeTreeSource)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(ole, nameof(ole));

                var newShape = new ShapeSc(innerTransform, spContext, ShapeContentType.OLEObject, shapeTreeSource)
                {
                    _ole = ole
                };

                return newShape;
            }

            public ShapeSc WithPicture(
                ILocation innerTransform, 
                ShapeContext spContext, 
                PictureSc picture, 
                GeometryType geometry, 
                OpenXmlCompositeElement shapeTreeSource)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(picture, nameof(picture));
                Check.NotNull(geometry, nameof(geometry));

                var newShape = new ShapeSc(innerTransform, spContext, ShapeContentType.Picture, geometry, shapeTreeSource)
                {
                    _picture = picture
                };

                return newShape;
            }

            public ShapeSc WithAutoShape(
                ILocation innerTransform, 
                ShapeContext spContext, 
                GeometryType geometry, 
                OpenXmlCompositeElement shapeTreeSource)
            {
                var newShape = new ShapeSc(innerTransform, spContext, ShapeContentType.AutoShape, geometry, shapeTreeSource);
      
                return newShape;
            }

            public ShapeSc WithTable(ILocation innerTransform, ShapeContext spContext, TableSc table, OpenXmlCompositeElement shapeTreeSource)
            {
                var newShape = new ShapeSc(innerTransform, spContext, ShapeContentType.Table, shapeTreeSource)
                {
                    _table = table
                };
                table.Shape = newShape;
                return newShape;
            }

            public ShapeSc WithChart(ILocation innerTransform, ShapeContext spContext, ChartSc chart, OpenXmlCompositeElement shapeTreeSource)
            {
                var newShape = new ShapeSc(innerTransform, spContext, ShapeContentType.Chart, shapeTreeSource)
                {
                    _chart = chart
                };

                return newShape;
            }

            public ShapeSc WithGroup(ILocation innerTransform, ShapeContext spContext, IList<ShapeSc> groupedShapes, OpenXmlCompositeElement shapeTreeSource)
            {
                var newShape = new ShapeSc(innerTransform, spContext, ShapeContentType.Group, shapeTreeSource)
                {
                    GroupedShapes = groupedShapes
                };

                return newShape;
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}