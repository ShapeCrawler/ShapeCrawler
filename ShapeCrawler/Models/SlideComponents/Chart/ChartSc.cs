using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Enums;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Settings;
using ShapeCrawler.Spreadsheet;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Models.SlideComponents.Chart
{
    public class ChartSc
    {
        #region Fields

        // Contains chart elements, e.g. <c:pieChart>. If the chart type is not a combination,
        // then collection contains only single item.
        private List<OpenXmlElement> _sdkCharts;

        private readonly ShapeContext _shapeContext;
        private readonly P.GraphicFrame _grFrame;
        private readonly Lazy<ChartType> _chartType;
        private readonly Lazy<OpenXmlElement> _firstSeries;
        private C.Chart _cChart;
        private Lazy<SeriesCollection> _seriesCollection;
        private Lazy<CategoryCollection> _categories;
        private readonly Lazy<LibraryCollection<double>> _xValues;
        private string _chartTitle;
        private ChartPart _sdkChartPart;
        private readonly ChartRefParser _chartRefParser;

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
                if (_chartTitle == null)
                {
                    _chartTitle = TryGetTitle();
                }

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
                if (_chartTitle == null)
                {
                    _chartTitle = TryGetTitle();
                }

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
#if NETSTANDARD2_1 || NETCOREAPP2_0
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

        #endregion Public Properties

        #region Constructors

        public ChartSc(P.GraphicFrame grFrame, ShapeContext shapeContext)
        : this(grFrame, shapeContext, new ChartRefParser(shapeContext))
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSc"/> class.
        /// </summary>
        public ChartSc(P.GraphicFrame grFrame, ShapeContext shapeContext, ChartRefParser chartRefParser)
        {
            _grFrame = grFrame ?? throw new ArgumentNullException(nameof(grFrame));
            _shapeContext = shapeContext ?? throw new ArgumentNullException(nameof(shapeContext));
            _chartRefParser = chartRefParser;
            _chartType = new Lazy<ChartType>(GetChartType);
            _firstSeries = new Lazy<OpenXmlElement>(GetFirstSeries);
            _xValues = new Lazy<LibraryCollection<double>>(TryGetXValues);
            Init(); //TODO: convert to lazy loading
        }

#endregion

        #region Private Methods

        private void Init()
        {
            var chartPartRef = _grFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>().GetFirstChild<C.ChartReference>().Id;
            _sdkChartPart = (ChartPart)_shapeContext.SdkSlidePart.GetPartById(chartPartRef);

            _cChart = _sdkChartPart.ChartSpace.GetFirstChild<C.Chart>();
            _sdkCharts = _cChart.PlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToList();  // example: <c:barChart>, <c:lineChart>
            _seriesCollection = new Lazy<SeriesCollection>(GetSeriesCollection);
            _categories = new Lazy<CategoryCollection>(TryGetCategories);
        }

        private ChartType GetChartType()
        {
            if (_sdkCharts.Count > 1)
            {
                return ChartType.Combination;
            }

            var chartName = _sdkCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType chartType);
            return chartType;
        }

        private SeriesCollection GetSeriesCollection()
        {
            return new SeriesCollection(_sdkCharts, _sdkChartPart, _chartRefParser);
        }

        private string TryGetTitle()
        {
            var title = _cChart.Title;
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

        private CategoryCollection TryGetCategories()
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
            var points = _chartRefParser.GetNumbers(sdkXValues.NumberReference, _sdkChartPart);

            return new LibraryCollection<double>(points);
        }

        private OpenXmlElement GetFirstSeries()
        {
            return _sdkCharts.First().ChildElements.FirstOrDefault(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
        }

        #endregion
    }
}


