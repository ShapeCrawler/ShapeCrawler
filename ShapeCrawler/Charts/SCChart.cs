using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    internal class SCChart : SlideShape, IChart
    {
        private readonly Lazy<ICategoryCollection> categories;
        private readonly Lazy<ChartType> chartType;
        private readonly Lazy<OpenXmlElement> firstSeries;
        private readonly P.GraphicFrame pGraphicFrame;
        private readonly Lazy<SeriesCollection> seriesCollection;
        private readonly Lazy<LibraryCollection<double>> xValues;
        private string? chartTitle;

        // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
        // then collection contains only single item.
        private IEnumerable<OpenXmlElement> cXCharts;

        internal SCChart(P.GraphicFrame pGraphicFrame, SCSlide parentSlideInternal)
            : base(pGraphicFrame, parentSlideInternal, null)
        { 
            this.pGraphicFrame = pGraphicFrame;
            this.firstSeries = new Lazy<OpenXmlElement>(this.GetFirstSeries);
            this.xValues = new Lazy<LibraryCollection<double>>(this.GetXValues);
            this.seriesCollection = new Lazy<SeriesCollection>(() => Collections.SeriesCollection.Create(this, this.cXCharts));
            this.categories = new Lazy<ICategoryCollection>(() => CategoryCollection.Create(this, this.firstSeries.Value, this.Type));
            this.chartType = new Lazy<ChartType>(this.GetChartType);
            this.ChartWorkbook = new ChartWorkbook(this);

            this.Init(); // TODO: convert to lazy loading
        }

        #region Public Properties

        public ChartType Type => this.chartType.Value;

        public string Title
        {
            get
            {
                this.chartTitle = this.GetTitleOrDefault();
                return this.chartTitle;
            }
        }

        public bool HasTitle
        {
            get
            {
                this.chartTitle ??= this.GetTitleOrDefault();
                return this.chartTitle != null;
            }
        }

        public bool HasCategories => categories.Value != null;

        public ISeriesCollection SeriesCollection => this.seriesCollection.Value;

        public ICategoryCollection Categories => this.categories.Value;

        public bool HasXValues => this.xValues.Value != null;

        public LibraryCollection<double> XValues
        {
            get
            {
                if (this.xValues.Value == null)
                {
                    throw new NotSupportedException(ExceptionMessages.NotXValues);
                }

                return this.xValues.Value;
            }
        }

        public override GeometryType GeometryType => GeometryType.Rectangle;

        public byte[] WorkbookByteArray => this.ChartWorkbook.ByteArray;

        #endregion Public Properties

        internal ChartWorkbook ChartWorkbook { get; }

        internal ChartPart SdkChartPart { get; private set; }

        private void Init()
        {
            // Get chart part
            C.ChartReference cChartReference = this.pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>()
                .GetFirstChild<C.ChartReference>();

            var slide = this.Slide;
            this.SdkChartPart = (ChartPart)slide.SlidePart.GetPartById(cChartReference.Id);

            C.PlotArea cPlotArea = this.SdkChartPart.ChartSpace.GetFirstChild<C.Chart>().PlotArea;
            this.cXCharts = cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        }

        private ChartType GetChartType()
        {
            if (this.cXCharts.Count() > 1)
            {
                return ChartType.Combination;
            }

            var chartName = this.cXCharts.Single().LocalName;
            Enum.TryParse(chartName, true, out ChartType enumChartType);

            return enumChartType;
        }

        private string GetTitleOrDefault()
        {
            C.Title cTitle = this.SdkChartPart.ChartSpace.GetFirstChild<C.Chart>().Title;
            if (cTitle == null)
            {
                // chart has not title
                return null;
            }

            C.ChartText cChartText = cTitle.ChartText;
            bool staticAvailable = this.TryGetStaticTitle(cChartText, out var staticTitle);
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
                return ((SeriesCollection) this.SeriesCollection).First().Name;
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
            var sdkXValues = firstSeries.Value?.GetFirstChild<C.XValues>();
            if (sdkXValues?.NumberReference == null)
            {
                return null;
            }

            IEnumerable<double> points =
                ChartReferencesParser.GetNumbersFromCacheOrWorkbook(sdkXValues.NumberReference, this);

            return new LibraryCollection<double>(points);
        }

        private OpenXmlElement GetFirstSeries()
        {
            return cXCharts.First().ChildElements
                .FirstOrDefault(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
        }

    }
}