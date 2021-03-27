using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart on a Slide.
    /// </summary>
    internal class SlideChart : SlideShape, IChart
    {
        private readonly Lazy<CategoryCollection> _categories;
        private readonly Lazy<ChartType> _chartType;
        private readonly Lazy<OpenXmlElement> _firstSeries;
        private readonly P.GraphicFrame _pGraphicFrame;
        private readonly Lazy<SeriesCollection> _seriesCollection;
        private readonly Lazy<LibraryCollection<double>> _xValues;
        private string _chartTitle;

        // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
        // then collection contains only single item.
        private IEnumerable<OpenXmlElement> _cXCharts;
        internal ChartPart ChartPart;

        #region Constructors

        /// <summary>
        ///     Initializes a new instance of the <see cref="SlideChart" /> class.
        /// </summary>
        internal SlideChart(P.GraphicFrame pGraphicFrame, SCSlide slide) :
            base(slide, pGraphicFrame)
        {
            _pGraphicFrame = pGraphicFrame;
            _firstSeries = new Lazy<OpenXmlElement>(GetFirstSeries);
            _xValues = new Lazy<LibraryCollection<double>>(GetXValues);
            _seriesCollection =
                new Lazy<SeriesCollection>(() => Collections.SeriesCollection.Create(this, _cXCharts));
            _categories = new Lazy<CategoryCollection>(() => CategoryCollection.Create(this, _firstSeries.Value, Type));
            _chartType = new Lazy<ChartType>(GetChartType);
            ChartWorkbook = new ChartWorkbook(this);

            Init(); //TODO: convert to lazy loading
        }

        #endregion Constructors

        internal ChartWorkbook ChartWorkbook { get; }

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
        public ISeriesCollection SeriesCollection => _seriesCollection.Value;

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

        public override GeometryType GeometryType => GeometryType.Rectangle;

        #endregion Public Properties

        #region Private Methods

        private void Init()
        {
            // Get chart part
            C.ChartReference cChartReference = _pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>()
                .GetFirstChild<C.ChartReference>();
            ChartPart = (ChartPart) Slide.SlidePart.GetPartById(cChartReference.Id);

            C.PlotArea cPlotArea = ChartPart.ChartSpace.GetFirstChild<C.Chart>().PlotArea;
            _cXCharts = cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
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
            C.Title cTitle = ChartPart.ChartSpace.GetFirstChild<C.Chart>().Title;
            if (cTitle == null) // chart has not title
            {
                return null;
            }

            C.ChartText cChartText = cTitle.ChartText;
            bool staticAvailable = TryGetStaticTitle(cChartText, out var staticTitle);
            if (staticAvailable)
            {
                return staticTitle;
            }

            // Dynamic title
            if (cChartText != null)
            {
                return cChartText.Descendants<C.StringPoint>().Single().InnerText;
            }

            // PieChart uses only one series for view.
            // However, it can have store multiple series data in the spreadsheet.
            if (Type == ChartType.PieChart)
            {
                return ((SeriesCollection) SeriesCollection).First().Name;
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

            IReadOnlyList<double> points =
                ChartReferencesParser.GetNumbersFromCacheOrSpreadsheet(sdkXValues.NumberReference, this);

            return new LibraryCollection<double>(points);
        }

        private OpenXmlElement GetFirstSeries()
        {
            return _cXCharts.First().ChildElements
                .FirstOrDefault(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
        }

        #endregion Private Methods
    }
}