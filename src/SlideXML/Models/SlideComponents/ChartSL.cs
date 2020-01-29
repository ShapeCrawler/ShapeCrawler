using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Enums;
using SlideXML.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a chart.
    /// </summary>
    public class ChartSL
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private ChartType? _type;
        private string _title;
        private OpenXmlElement _xmlElement; // contains chart element, e.g. <c:pieChart>
        private C.Chart _cChart;

        #endregion

        #region Properties

        /// <summary>
        /// Returns the chart type.
        /// </summary>
        public ChartType Type
        {
            get
            {
                if (_type == null)
                {
                    ParseType();
                }

                return (ChartType)_type;
            }
        }

        /// <summary>
        /// Returns the chart title text.
        /// </summary>
        public string Title => _title ??= ParseTitle();

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartSL"/> class.
        /// </summary>
        public ChartSL(P.GraphicFrame grFrame, SlidePart sldPart)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            _sldPart = sldPart;
            _xmlElement = grFrame;

            Init();
        }

        #endregion

        private void Init()
        {
            // Get reference
            var chartRef = _xmlElement.Descendants<C.ChartReference>().Single();

            // Get chart part by reference
            var chPart = _sldPart.GetPartById(chartRef.Id) as ChartPart;

            _cChart = chPart.ChartSpace.GetFirstChild<C.Chart>();
            _xmlElement = _cChart.PlotArea.Elements().Single(e => e.LocalName.EndsWith("Chart"));
        }

        private void ParseType()
        {
            var chartName = _xmlElement.LocalName;
            var parsed = Enum.TryParse(chartName, true, out ChartType chartType);
            if (!parsed)
            {
                throw new SlideXMLException("An error occured during parse chart type.");
            }
            _type = chartType;
        }

        private string ParseTitle()
        {
            var chartText = _cChart.Title.ChartText;

            // First, try parse static title
            var rRich = chartText?.RichText;
            if (rRich != null)
            {
                return rRich.Descendants<A.Text>().Single().Text;
            }

            // Dynamic title
            if (chartText != null)
            {
                return chartText.Descendants<C.StringPoint>().Single().InnerText;
            }
            // Parse PieChart dynamic title
            return _xmlElement.GetFirstChild<C.PieChartSeries>().GetFirstChild<C.SeriesText>().Descendants<C.StringPoint>().Single().InnerText;
        }
    }
}



