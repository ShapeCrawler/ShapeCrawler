using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Enums;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Models.SlideComponents.Chart;
using ShapeCrawler.Models.TextBody;
using ShapeCrawler.Models.Transforms;
using ShapeCrawler.Services.Builders;
using ShapeCrawler.Services.Drawing;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using SlideDotNet.Models.TableComponents;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    /// Represents a shape on a slide.
    /// </summary>
    public class ShapeEx
    {
        #region Fields

        private readonly IShapeContext _context;
        private readonly Lazy<ITextFrame> _textFrame;
        private readonly Lazy<Fill> _fill;
        private readonly IImageExFactory _imageFactory = new ImageExFactory(); //TODO: do not initiate for non-AutoShape types
        private bool? _hidden;
        private int _id;
        private string _name;
        private PictureEx _picture;
        private OleObject _ole;
        private TableEx _table;
        private ChartEx _chart;
        private readonly ILocation _innerTransform;

        #endregion Fields

        #region Properties

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
        public bool HasTextFrame => TextFrame is TextFrame;

        /// <summary>
        /// Determines whether the shape has chart content.
        /// </summary>
        public bool HasChart => Chart is Chart.ChartEx;

        /// <summary>
        /// Determines whether the slide element has picture content.
        /// </summary>
        public bool HasPicture => _picture != null;

        /// <summary>
        /// Returns text frame.
        /// </summary>
        /// <remarks>Lazy load.</remarks>
        public ITextFrame TextFrame => _textFrame.Value;

        /// <summary>
        /// Returns chart. Throws exception if shape content type is not <see cref="ShapeContentType.Chart"/>
        /// </summary>
        public IChart Chart => _chart ?? throw new NotSupportedException(ExceptionMessages.NoChart);

        /// <summary>
        /// Returns table. Throws exception if shape content type is not <see cref="ShapeContentType.Table"/>
        /// </summary>
        public TableEx Table => _table ?? throw new NotSupportedException(ExceptionMessages.NoTable);

        /// <summary>
        /// Returns picture. Throws exception if shape content type is not a <see cref="ShapeContentType.Picture"/>
        /// </summary>
        public PictureEx Picture => _picture ?? throw new NotSupportedException(ExceptionMessages.NoPicture);

        /// <summary>
        /// Returns grouped shapes. Throws exception if shape content type is not <see cref="ShapeContentType.Group"/>
        /// </summary>
        public IList<ShapeEx> GroupedShapes { get; private set; }

        /// <summary>
        /// Returns OLE object content.
        /// </summary>
        public OleObject OleObject => _ole ?? throw new NotSupportedException(ExceptionMessages.NoOleObject);

        /// <summary>
        /// Determines whether the shape is placeholder.
        /// </summary>
        public bool IsPlaceholder => _context.SdkElement.IsPlaceholder();

        /// <summary>
        /// Returns placeholder type. Returns null if shape is not a placeholder.
        /// </summary>
        public PlaceholderType PlaceholderType
        {
            get
            {
                if (!IsPlaceholder)
                {
                    throw new NotSupportedException(ExceptionMessages.ShapeIsNotPlaceholder);
                }
                
                return _context.PlaceholderService.GetPlaceholderType(_context.SdkElement);
            }
        }

        /// <summary>
        /// Returns shape fill. Returns <c>null</c> if shape has not fill.
        /// </summary>
        public Fill Fill => _fill.Value;

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

        private void SetCustomData(string value)
        {
            var customDataElement = $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            _context.SdkElement.InnerXml += customDataElement;
        }

        #endregion Properties

        #region Constructors

        private ShapeEx(ILocation innerTransform,
                        IShapeContext spContext,
                        ShapeContentType contentType,
                        GeometryType geometryType) : this(innerTransform, spContext, contentType)
        {
            GeometryType = geometryType;
        }

        private ShapeEx (ILocation innerTransform, IShapeContext spContext, ShapeContentType contentType)
        {
            _innerTransform = innerTransform;
            _context = spContext;
            ContentType = contentType;
            _textFrame = new Lazy<ITextFrame>(TryGetTextFrame);
            _fill = new Lazy<Fill>(TryGetFill);
        }

        #endregion Constructors

        #region Private Methods

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(_context.SdkElement.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        private ITextFrame TryGetTextFrame()
        {
            if (ContentType != ShapeContentType.AutoShape)
            {
                return new NoTextFrame();
            }

            var pTxtBody = _context.SdkElement.Descendants<P.TextBody>().SingleOrDefault();
            if (pTxtBody == null)
            {
                return new NoTextFrame();
            }

            var aTexts = pTxtBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element with text must be exist
            {
                return new TextFrame(_context, pTxtBody);
            }

            return new NoTextFrame();
        }

        private Fill TryGetFill()
        {
            if (ContentType != ShapeContentType.AutoShape)
            {
                return null;
            }
            var image = _imageFactory.TryFromSdkShape(_context.SkdSlidePart, (OpenXmlCompositeElement)_context.SdkElement); //TODO: delete casting
            if (image != null)
            {
                return new Fill(image);
            }

            var xmlShape = (P.Shape) _context.SdkElement;
            var rgbColorModelHex = xmlShape.ShapeProperties.GetFirstChild<A.SolidFill>()?.RgbColorModelHex;
            if (rgbColorModelHex != null)
            {
                return Fill.FromXmlSolidFill(rgbColorModelHex);
            }

            return null;
        }

        private void InitIdHiddenName()
        {
            if (_id != 0)
            {
                return;
            }
            var (id, hidden, name) = ((OpenXmlCompositeElement)_context.SdkElement).GetNvPrValues(); //TODO: delete casting
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

            public ShapeEx WithOle(ILocation innerTransform, IShapeContext spContext, OleObject ole)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(ole, nameof(ole));

                var newShape = new ShapeEx(innerTransform, spContext, ShapeContentType.OLEObject)
                {
                    _ole = ole
                };

                return newShape;
            }

            public ShapeEx WithPicture(ILocation innerTransform, IShapeContext spContext, PictureEx picture, GeometryType geometry)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(picture, nameof(picture));
                Check.NotNull(geometry, nameof(geometry));

                var newShape = new ShapeEx(innerTransform, spContext, ShapeContentType.Picture, geometry)
                {
                    _picture = picture
                };

                return newShape;
            }

            public ShapeEx WithAutoShape(ILocation innerTransform, IShapeContext spContext, GeometryType geometry)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));

                var newShape = new ShapeEx(innerTransform, spContext, ShapeContentType.AutoShape, geometry);
      
                return newShape;
            }

            public ShapeEx WithTable(ILocation innerTransform, IShapeContext spContext, TableEx table)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(table, nameof(table));

                var newShape = new ShapeEx(innerTransform, spContext, ShapeContentType.Table)
                {
                    _table = table
                };

                return newShape;
            }

            public ShapeEx WithChart(ILocation innerTransform, IShapeContext spContext, ChartEx chart)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(chart, nameof(chart));

                var newShape = new ShapeEx(innerTransform, spContext, ShapeContentType.Chart)
                {
                    _chart = chart
                };

                return newShape;
            }

            public ShapeEx WithGroup(ILocation innerTransform, IShapeContext spContext, IList<ShapeEx> groupedShapes)
            {
                Check.NotNull(innerTransform, nameof(innerTransform));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(groupedShapes, nameof(groupedShapes));

                var newShape = new ShapeEx(innerTransform, spContext, ShapeContentType.Group)
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