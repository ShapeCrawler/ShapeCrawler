using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using ShapeCrawler.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    /// <inheritdoc cref="IChart" />
    public class SlideChart : SlideShape, IChart
    {
        #region Constructors

        /// <summary>
        ///     Initializes a new instance of the <see cref="SlideChart" /> class.
        /// </summary>
        internal SlideChart(
            P.GraphicFrame pGraphicFrame,
            SlideSc slide,
            ILocation innerTransform,
            ShapeContext spContext) : base(slide, pGraphicFrame)
        {
            _pGraphicFrame = pGraphicFrame;
            _innerTransform = innerTransform;
            Context = spContext;

            _firstSeries = new Lazy<OpenXmlElement>(GetFirstSeries);
            _xValues = new Lazy<LibraryCollection<double>>(GetXValues);
            _seriesCollection =
                new Lazy<SeriesCollection>(() => SeriesCollection.Create(_cXCharts, _chartPart, _chartRefParser));
            _categories = new Lazy<CategoryCollection>(() => CategoryCollection.Create(_firstSeries.Value, Type));
            _chartRefParser = new ChartReferencesParser(this);
            _chartType = new Lazy<ChartType>(GetChartType);

            Init(); //TODO: convert to lazy loading
        }

        #endregion Constructors

        #region Private Methods

        private void
            InitIdHiddenName() // TODO: check, looks like it can be shared and can be moved int base Shape class.
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

        #region Fields

        // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
        // then collection contains only single item.
        private IEnumerable<OpenXmlElement> _cXCharts;

        private bool? _hidden;
        private int _id;
        private string _name;
        private readonly ILocation _innerTransform;
        private readonly Lazy<ChartType> _chartType;
        private readonly Lazy<OpenXmlElement> _firstSeries;
        private readonly Lazy<SeriesCollection> _seriesCollection;
        private readonly Lazy<CategoryCollection> _categories;
        private readonly Lazy<LibraryCollection<double>> _xValues;
        private string _chartTitle;
        private ChartPart _chartPart;
        private readonly ChartReferencesParser _chartRefParser;
        private readonly P.GraphicFrame _pGraphicFrame;

        internal ShapeContext Context { get; }

        #endregion Fields

        #region Public Properties

        /// <summary>
        ///     Gets the chart title. Returns null if chart has not a title.
        /// </summary>
        public ChartType Type => _chartType.Value;

        /// <summary>
        ///     Gets chart title string.
        /// </summary>
        public string Title
        {
            get
            {
                _chartTitle ??= TryGetTitle();

                return _chartTitle ?? throw new NotSupportedException(ExceptionMessages.NotTitle);
            }
        }

        /// <summary>
        ///     Determines whether chart has a title.
        /// </summary>
        public bool HasTitle
        {
            get
            {
                _chartTitle ??= TryGetTitle();

                return _chartTitle != null;
            }
        }

        /// <summary>
        ///     Determines whether chart has categories. Some chart types like ScatterChart and BubbleChart does not have
        ///     categories.
        /// </summary>
        public bool HasCategories => _categories.Value != null;

        /// <summary>
        ///     Gets collection of the chart series.
        /// </summary>
        public SeriesCollection SeriesCollection => _seriesCollection.Value;

        /// <summary>
        ///     Gets chart categories. Returns <c>NULL</c> if the chart does not have categories.
        /// </summary>
        public CategoryCollection Categories => _categories.Value;

        public bool HasXValues => _xValues.Value != null;

        public LibraryCollection<double> XValues
        {
            get
            {
                if (_xValues.Value == null)
                {
                    throw new NotSupportedException(ExceptionMessages.NotXValues);
                }

                return _xValues.Value;
            }
        }

        #endregion Public Properties

        #region Private Methods

        private void Init()
        {
            StringValue chartPartRef = _pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>()
                .GetFirstChild<C.ChartReference>().Id;
            _chartPart = (ChartPart) Slide.SlidePart.GetPartById(chartPartRef);
            _cXCharts = _chartPart.ChartSpace.GetFirstChild<C.Chart>().PlotArea
                .Where(e => e.LocalName.EndsWith("Chart",
                    StringComparison.Ordinal));
        }

        private ChartType GetChartType()
        {
            if (_cXCharts.Count() > 1)
            {
                return ChartType.Combination;
            }

            string chartName = _cXCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType chartType);
            return chartType;
        }

        private string TryGetTitle()
        {
            var title = _chartPart.ChartSpace.GetFirstChild<C.Chart>().Title;
            if (title == null) // chart has not title
            {
                return null;
            }

            var xmlChartText = title.ChartText;
            var staticAvailable = TryGetStaticTitle(xmlChartText, out var staticTitle);
            if (staticAvailable)
            {
                return staticTitle;
            }

            // Dynamic title
            if (xmlChartText != null)
            {
                return xmlChartText.Descendants<C.StringPoint>().Single().InnerText;
            }

            // PieChart uses only one series for view.
            // However, it can have store multiple series data in the spreadsheet.
            if (Type == ChartType.PieChart)
            {
                return SeriesCollection.First().Name;
            }

            return null;
        }

        private bool TryGetStaticTitle(C.ChartText chartText, out string staticTitle)
        {
            staticTitle = null;
            if (Type == ChartType.Combination)
            {
                staticTitle = chartText.RichText.Descendants<A.Text>().Select(t => t.Text)
                    .Aggregate((t1, t2) => t1 + t2);
                return true;
            }

            var rRich = chartText?.RichText;
            if (rRich != null)
            {
                staticTitle = rRich.Descendants<A.Text>().Select(t => t.Text).Aggregate((t1, t2) => t1 + t2);
                return true;
            }

            return false;
        }

        private LibraryCollection<double> GetXValues()
        {
            var sdkXValues = _firstSeries.Value?.GetFirstChild<C.XValues>();
            if (sdkXValues?.NumberReference == null)
            {
                return null;
            }

            var points = _chartRefParser.GetNumbersFromCacheOrSpreadsheet(sdkXValues.NumberReference, _chartPart);

            return new LibraryCollection<double>(points);
        }

        private OpenXmlElement GetFirstSeries()
        {
            return _cXCharts.First().ChildElements
                .FirstOrDefault(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
        }

        #endregion Private Methods

        #region Public Properties

        /// <summary>
        ///     Returns the x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X
        {
            get => _innerTransform.X;
            set => _innerTransform.SetX(value);
        }

        /// <summary>
        ///     Returns the y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y
        {
            get => _innerTransform.Y;
            set => _innerTransform.SetY(value);
        }

        /// <summary>
        ///     Returns the width of the shape.
        /// </summary>
        public long Width
        {
            get => _innerTransform.Width;
            set => _innerTransform.SetWidth(value);
        }

        /// <summary>
        ///     Returns the height of the shape.
        /// </summary>
        public long Height
        {
            get => _innerTransform.Height;
            set => _innerTransform.SetHeight(value);
        }

        /// <summary>
        ///     Returns an element identifier.
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
        ///     Gets an element name.
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
        ///     Determines whether the shape is hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                InitIdHiddenName();
                return (bool) _hidden;
            }
        }

        public override GeometryType GeometryType => GeometryType.Rectangle;

        #endregion Properties
    }
}