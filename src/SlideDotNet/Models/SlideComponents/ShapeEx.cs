using System;
using System.Collections.Generic;
using System.Linq;
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
using SlideXML.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models.SlideComponents
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
        private Picture _picture;
        private OleObject _ole;
        private TableEx _table;
        private ChartEx _chart;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns the x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X { get; set; } //TODO: delete public setter

        /// <summary>
        /// Returns the y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y { get; set; } //TODO: delete public setter

        /// <summary>
        /// Returns the width of the shape.
        /// </summary>
        public long Width { get; }

        /// <summary>
        /// Returns the height of the shape.
        /// </summary>
        public long Height { get; }

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
        public IChart Chart => _chart ?? throw new SlideDotNetException(ExceptionMessages.NoChart);

        /// <summary>
        /// Returns table. Throws exception if shape content type is not <see cref="ShapeContentType.Table"/>
        /// </summary>
        public TableEx Table => _table ?? throw new SlideDotNetException(ExceptionMessages.NoTable);

        /// <summary>
        /// Returns picture. Throws exception if shape content type is not a <see cref="ShapeContentType.Picture"/>
        /// </summary>
        public Picture Picture => _picture ?? throw new SlideDotNetException(ExceptionMessages.NoPicture);

        /// <summary>
        /// Returns grouped shapes. Throws exception if shape content type is not <see cref="ShapeContentType.Group"/>
        /// </summary>
        public IList<ShapeEx> GroupedShapes { get; private set; }

        /// <summary>
        /// Returns OLE object content.
        /// </summary>
        public OleObject OleObject => _ole ?? throw new SlideDotNetException(ExceptionMessages.NoOleObject);

        /// <summary>
        /// Returns placeholder type. Returns null if shape is not a placeholder.
        /// </summary>
        public PlaceholderType? PlaceholderType
        {
            get
            {
                if (_context.XmlElement.IsPlaceholder())
                {
                    return PlaceholderService.PlaceholderDataFrom(_context.XmlElement).PlaceholderType;
                }

                return null;
            }
        }

        /// <summary>
        /// Returns shape fill. Returns <c>null</c> if shape has not fill.
        /// </summary>
        public Fill Fill => _fill.Value;

        #endregion Properties

        #region Constructors

        private ShapeEx (Location location, IShapeContext spContext, ShapeContentType contentType)
        {
            X = location.X;
            Y = location.Y;
            Width = location.Width;
            Height = location.Height;
            _context = spContext;
            ContentType = contentType;
            _textFrame = new Lazy<ITextFrame>(TryGetTextFrame);
            _fill = new Lazy<Fill>(TryGetFill);
        }

        #endregion Constructors

        #region Private Methods

        private ITextFrame TryGetTextFrame()
        {
            if (ContentType != ShapeContentType.AutoShape)
            {
                return new NoTextFrame();
            }

            var pTxtBody = _context.XmlElement.Descendants<P.TextBody>().SingleOrDefault();
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
            var image = _imageFactory.TryFromXmlShape(_context.XmlSlidePart, _context.XmlElement);
            if (image != null)
            {
                return new Fill(image);
            }

            var xmlShape = (P.Shape) _context.XmlElement;
            var solidFill = xmlShape.ShapeProperties.GetFirstChild<A.SolidFill>();
            if (solidFill != null)
            {
                return Fill.FromXmlSolidFill(solidFill);
            }

            return null;
        }

        private void InitIdHiddenName()
        {
            if (_id != 0)
            {
                return;
            }
            var (id, hidden, name) = _context.XmlElement.GetNvPrValues();
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

            /// <summary>
            /// <inheritdoc cref="IShapeBuilder.WithOle"/>
            /// </summary>
            public ShapeEx WithOle(Location location, IShapeContext spContext, OleObject ole)
            {
                Check.NotNull(location, nameof(location));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(ole, nameof(ole));

                var newShape = new ShapeEx(location, spContext, ShapeContentType.OLEObject)
                {
                    _ole = ole
                };

                return newShape;
            }

            /// <summary>
            /// <inheritdoc cref="IShapeBuilder.WithPicture"/>
            /// </summary>
            public ShapeEx WithPicture(Location location, IShapeContext spContext, Picture picture)
            {
                Check.NotNull(location, nameof(location));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(picture, nameof(picture));

                var newShape = new ShapeEx(location, spContext, ShapeContentType.Picture)
                {
                    _picture = picture
                };

                return newShape;
            }

            /// <summary>
            /// <inheritdoc cref="IShapeBuilder.WithAutoShape"/>
            /// </summary>
            public ShapeEx WithAutoShape(Location location, IShapeContext spContext)
            {
                Check.NotNull(location, nameof(location));
                Check.NotNull(spContext, nameof(spContext));

                var newShape = new ShapeEx(location, spContext, ShapeContentType.AutoShape);
      
                return newShape;
            }

            /// <summary>
            /// <inheritdoc cref="IShapeBuilder.WithTable"/>
            /// </summary>
            public ShapeEx WithTable(Location location, IShapeContext spContext, TableEx table)
            {
                Check.NotNull(location, nameof(location));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(table, nameof(table));

                var newShape = new ShapeEx(location, spContext, ShapeContentType.Table)
                {
                    _table = table
                };

                return newShape;
            }

            /// <summary>
            /// <inheritdoc cref="IShapeBuilder.WithChart"/>
            /// </summary>
            public ShapeEx WithChart(Location location, IShapeContext spContext, ChartEx chart)
            {
                Check.NotNull(location, nameof(location));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(chart, nameof(chart));

                var newShape = new ShapeEx(location, spContext, ShapeContentType.Chart)
                {
                    _chart = chart
                };

                return newShape;
            }

            /// <summary>
            /// <inheritdoc cref="IShapeBuilder.WithGroup"/>
            /// </summary>
            public ShapeEx WithGroup(Location location, IShapeContext spContext, IEnumerable<ShapeEx> groupedShapes)
            {
                Check.NotNull(location, nameof(location));
                Check.NotNull(spContext, nameof(spContext));
                Check.NotNull(groupedShapes, nameof(groupedShapes));

                var newShape = new ShapeEx(location, spContext, ShapeContentType.Group)
                {
                    GroupedShapes = groupedShapes.ToList()
                };

                return newShape;
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}