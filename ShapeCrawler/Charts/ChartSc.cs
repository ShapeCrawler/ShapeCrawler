using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using ShapeCrawler.Spreadsheet;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{

    /// <summary>
    /// Represents a shape on a slide.
    /// </summary>
    public class ChartSc : IChart
    {
        #region Fields

        // Contains chart elements, e.g. <c:pieChart>. If the chart type is not a combination,
        // then collection contains only single item.
        private IEnumerable<OpenXmlElement> _sdkCharts;

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
        private readonly ChartRefParser _chartRefParser;
        private readonly P.GraphicFrame _pGraphicFrame;

        internal ShapeContext Context { get; }
        internal OpenXmlCompositeElement ShapeTreeSource { get; }
        internal SlideSc Slide { get; }

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

        #endregion Public Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSc"/> class.
        /// </summary>
        internal ChartSc(
            P.GraphicFrame pGraphicFrame, 
            SlideSc slide, 
            OpenXmlCompositeElement shapeTreeSource,
            ILocation innerTransform,
            ShapeContext spContext)
        {
            _pGraphicFrame = pGraphicFrame;
            Slide = slide;
            ShapeTreeSource = shapeTreeSource;
            _innerTransform = innerTransform;
            Context = spContext;

            _firstSeries = new Lazy<OpenXmlElement>(GetFirstSeries);
            _xValues = new Lazy<LibraryCollection<double>>(TryGetXValues);
            _seriesCollection = new Lazy<SeriesCollection>(GetSeriesCollection);
            _categories = new Lazy<CategoryCollection>(TryGetCategoryCollection);
            _chartRefParser = new ChartRefParser(this);
            _chartType = new Lazy<ChartType>(GetChartType);

            Init(); //TODO: convert to lazy loading
        }

        #endregion Constructors

        #region Private Methods

        private void Init()
        {
            StringValue chartPartRef = _pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>().GetFirstChild<C.ChartReference>().Id;
            _chartPart = Slide.SlidePart.GetPartById(chartPartRef) as ChartPart;
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

        #endregion Private Methods


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
        
        public GeometryType GeometryType { get; }

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }


        #endregion Properties

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

        private void InitIdHiddenName() // TODO: check, looks like it can be shared
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
    }
}