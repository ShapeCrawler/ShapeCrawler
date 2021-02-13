using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Spreadsheet;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    public class ChartSc
    {
        #region Fields

        // Contains chart elements, e.g. <c:pieChart>. If the chart type is not a combination,
        // then collection contains only single item.
        private IEnumerable<OpenXmlElement> _sdkCharts;

        private readonly Lazy<ChartType> _chartType;
        private readonly Lazy<OpenXmlElement> _firstSeries;
        private readonly Lazy<SeriesCollection> _seriesCollection;
        private readonly Lazy<CategoryCollection> _categories;
        private readonly Lazy<LibraryCollection<double>> _xValues;
        private string _chartTitle;
        private ChartPart _chartPart;
        private readonly ChartRefParser _chartRefParser;
        private readonly GraphicFrame _pGraphicFrame;
        private readonly SlideSc _slide;

        #endregion Fields

        #region Public Properties

        /// <summary>
        /// Gets the chart title. Returns null if chart has not a title.
        /// </summary>
        public ChartType Type => _chartType.Value;

        /// <summary>
        /// Gets chart title string.
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
        /// Determines whether chart has a title.
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
        /// Determines whether chart has categories. Some chart types like ScatterChart and BubbleChart does not have categories.
        /// </summary>
        public bool HasCategories => _categories.Value != null;

        /// <summary>
        /// Gets collection of the chart series.
        /// </summary>
        public SeriesCollection SeriesCollection => _seriesCollection.Value;

        /// <summary>
        /// Gets collection of the chart category.
        /// </summary>
        public CategoryCollection Categories
        {
            get
            {
                if (_categories.Value == null)
                {
#if NETSTANDARD2_1 || NETCOREAPP2_0 || NET5_0
                    var msg = ExceptionMessages.ChartCanNotHaveCategory.Replace("#0", Type.ToString(), StringComparison.OrdinalIgnoreCase);
#else
                    var msg = ExceptionMessages.ChartCanNotHaveCategory.Replace("#0", Type.ToString());
#endif
                    throw new NotSupportedException(msg);
                }

                return _categories.Value;
            }
        }

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

        internal ShapeSc Shape { get; set; }

        #endregion Public Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSc"/> class.
        /// </summary>
        internal ChartSc(P.GraphicFrame pGraphicFrame,  SlideSc slide)
        {
            _pGraphicFrame = pGraphicFrame;
            _chartRefParser = new ChartRefParser(this);
            _chartType = new Lazy<ChartType>(GetChartType);
            _firstSeries = new Lazy<OpenXmlElement>(GetFirstSeries);
            _xValues = new Lazy<LibraryCollection<double>>(TryGetXValues);
            _slide = slide;
            _seriesCollection = new Lazy<SeriesCollection>(GetSeriesCollection);
            _categories = new Lazy<CategoryCollection>(TryGetCategoryCollection);
            Init(); //TODO: convert to lazy loading
        }

        #endregion

        #region Private Methods

        private void Init()
        {
            StringValue chartPartRef = _pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>().GetFirstChild<C.ChartReference>().Id;
            _chartPart = _slide.SlidePart.GetPartById(chartPartRef) as ChartPart;
            _sdkCharts = _chartPart.ChartSpace.GetFirstChild<C.Chart>().PlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));  // example: <c:barChart>, <c:lineChart>
        }

        private ChartType GetChartType()
        {
            if (_sdkCharts.Count() > 1)
            {
                return ChartType.Combination;
            }

            var chartName = _sdkCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType chartType);
            return chartType;
        }

        private SeriesCollection GetSeriesCollection()
        {
            return new SeriesCollection(_sdkCharts, _chartPart, _chartRefParser);
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
                staticTitle = chartText.RichText.Descendants<A.Text>().Select(t => t.Text).Aggregate((t1, t2) => t1 + t2);
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

        private CategoryCollection TryGetCategoryCollection()
        {
            if (Type == ChartType.BubbleChart || Type == ChartType.ScatterChart)
            {
                return null;
            }
            
            return new CategoryCollection(_firstSeries.Value);
        }

        private LibraryCollection<double> TryGetXValues()
        {
            var sdkXValues = _firstSeries.Value?.GetFirstChild<C.XValues>();
            if (sdkXValues?.NumberReference == null)
            {
                return null;
            }
            var points = _chartRefParser.GetNumbers(sdkXValues.NumberReference, _chartPart);

            return new LibraryCollection<double>(points);
        }

        private OpenXmlElement GetFirstSeries()
        {
            return _sdkCharts.First().ChildElements.FirstOrDefault(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
        }

        #endregion
    }
}


