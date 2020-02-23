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
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using TableEx = SlideXML.Models.SlideComponents.TableEx;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models.SlideComponents
{
    /// <summary>
    /// Represents a slide element.
    /// </summary>
    public class ShapeEx
    {
        #region Fields

        private readonly IShapeContext _context;
        private readonly Lazy<ITextFrame> _textFrame;
        private readonly IBackgroundImageFactory _bgImgFactory = new BackgroundImageFactory(); //TODO: do not initiate for non-AutoShape types
        private bool? _hidden;
        private int _id;
        private string _name;
        private ImageEx _backgroundImage;
        private PlaceholderType? _placeholderType;
        private Picture _picture;
        private OleObject _ole;
        private TableEx _table;
        private ChartEx _chart;

        #endregion Fields

        #region Properties

        public long X { get; set; } //TODO: delete public setter

        public long Y { get; set; } //TODO: delete public setter

        public long Width { get; private set; }

        public long Height { get; private set; }

        public ElementType Type { get; private set; }

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
                return _backgroundImage ??= _bgImgFactory.FromXmlShape(_context.XmlSlidePart, (P.Shape) _context.XmlElement);
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
        public bool HasChart => Chart is Chart.ChartEx;

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
        public IChart Chart => _chart ?? throw new SlideDotNetException(ExceptionMessages.NoChart);

        /// <summary>
        /// Returns table.
        /// </summary>
        public TableEx Table => _table ?? throw new SlideDotNetException(ExceptionMessages.NoTable);

        /// <summary>
        /// Returns picture.
        /// </summary>
        public Picture Picture => _picture ?? throw new SlideDotNetException(ExceptionMessages.NoPicture);

        /// <summary>
        /// Returns grouped shapes. Throws exception if shape type is not <see cref="ElementType.Group"/>
        /// </summary>
        public IList<ShapeEx> GroupedShapes { get; private set; }

        /// <summary>
        /// Returns OLE object.
        /// </summary>
        public OleObject OleObject => _ole ?? throw new SlideDotNetException(ExceptionMessages.NoOleObject);

        /// <summary>
        /// Determines shape is placeholder.
        /// </summary>
        public bool IsPlaceholder => _context.XmlElement.IsPlaceholder(); //TODO: make lazy

        public PlaceholderType PlaceholderType
        {
            get
            {
                if (_placeholderType == null && _context.XmlElement.IsPlaceholder())
                {
                    _placeholderType = PlaceholderService.PlaceholderDataFrom(_context.XmlElement).PlaceholderType;
                }
                if (_placeholderType == null)
                {
                    throw new SlideDotNetException(ExceptionMessages.NotPlaceholder);
                }

                return (PlaceholderType)_placeholderType;
            }
        }

        #endregion Properties

        #region Constructors

        private ShapeEx (Location location, IShapeContext spContext, ElementType type)
        {
            X = location.X;
            Y = location.Y;
            Width = location.Width;
            Height = location.Height;
            _context = spContext;
            Type = type;
            _textFrame = new Lazy<ITextFrame>(GetTextFrame);
        }

        #endregion Constructors

        #region Private Methods

        private ITextFrame GetTextFrame()
        {
            if (Type != ElementType.AutoShape)
            {
                return new NoTextFrame();
            }

            var pTxtBody = _context.XmlElement.Descendants<P.TextBody>().SingleOrDefault();
            if (pTxtBody == null)
            {
                return new NoTextFrame();
            }

            var aTexts = pTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                return new TextFrame(_context, pTxtBody);
            }

            return new NoTextFrame();
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

                var newShape = new ShapeEx(location, spContext, ElementType.OLEObject)
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

                var newShape = new ShapeEx(location, spContext, ElementType.Picture)
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

                var newShape = new ShapeEx(location, spContext, ElementType.AutoShape);
      
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

                var newShape = new ShapeEx(location, spContext, ElementType.Table)
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

                var newShape = new ShapeEx(location, spContext, ElementType.Chart)
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

                var newShape = new ShapeEx(location, spContext, ElementType.Group)
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