using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    internal class SCChart : SlideShape, IChart
    {
        private readonly ResettableLazy<ICategoryCollection?> categories;
        private readonly Lazy<ChartType> chartType;
        private readonly Lazy<OpenXmlElement> firstSeries;
        private readonly P.GraphicFrame pGraphicFrame;
        private readonly Lazy<SCSeriesCollection> series;
        private readonly Lazy<LibraryCollection<double>> xValues;
        private string? chartTitle;

        // Contains chart elements, e.g. <c:pieChart>, <c:barChart>, <c:lineChart> etc. If the chart type is not a combination,
        // then collection contains only single item.
        private IEnumerable<OpenXmlElement> cXCharts;

        internal SCChart(P.GraphicFrame pGraphicFrame, SCSlide slideInternal)
            : base(pGraphicFrame, slideInternal, null)
        {
            this.pGraphicFrame = pGraphicFrame;
            this.firstSeries = new Lazy<OpenXmlElement>(this.GetFirstSeries);
            this.xValues = new Lazy<LibraryCollection<double>>(this.GetXValues);
            this.series = new Lazy<SCSeriesCollection>(this.GetSeries);
            this.categories = new ResettableLazy<ICategoryCollection?>(this.GetCategories);
            this.chartType = new Lazy<ChartType>(this.GetChartType);

            var cChartReference = this.pGraphicFrame.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !
                .GetFirstChild<C.ChartReference>();

            this.ChartPart = (ChartPart)this.Slide.SDKSlidePart.GetPartById(cChartReference.Id);
            
            var cPlotArea = this.ChartPart.ChartSpace.GetFirstChild<C.Chart>().PlotArea;
            this.cXCharts = cPlotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

            this.ChartWorkbook = this.ChartPart.EmbeddedPackagePart != null ? new ChartWorkbook(this, this.ChartPart.EmbeddedPackagePart) : null;
        }

        public ChartType Type => this.chartType.Value;

        public ShapeType ShapeType => ShapeType.Chart;

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

        public bool HasCategories => this.categories.Value != null;

        public ISeriesCollection SeriesCollection => this.series.Value;

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

        public SpreadsheetDocument SDKSpreadsheetDocument => this.ChartWorkbook.spreadsheetDocument.Value;

        internal ChartWorkbook? ChartWorkbook { get; set; }

        internal ChartPart ChartPart { get; private set; }

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
        
        private ICategoryCollection? GetCategories()
        {
            return CategoryCollection.Create(this, this.firstSeries.Value, this.Type);
        }

        private SCSeriesCollection GetSeries()
        {
            return SCSeriesCollection.Create(this, this.cXCharts);
        }

        private string GetTitleOrDefault()
        {
            C.Title cTitle = this.ChartPart.ChartSpace.GetFirstChild<C.Chart>().Title;
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
                return ((SCSeriesCollection) this.SeriesCollection).First().Name;
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