using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Exceptions;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.SlideComponents.Chart
{
    /// <summary>
    /// <inheritdoc cref="IChart"/>
    /// </summary>
    public class ChartEx : IChart
    {
        #region Fields

        // Contains chart elements, e.g. <c:pieChart>. If the chart type is not a combination,
        // then collection contains an only single item.
        private List<OpenXmlElement> _chartElements;
        private readonly SlidePart _sldPart;
        private ChartType? _type; //TODO: make lazy
        private string _title;
        private readonly P.GraphicFrame _grFrame;
        private C.Chart _cChart;

        #endregion Fields

        #region Properties

        /// <summary>
        /// <inheritdoc cref="IChart.Type"/>
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
        /// <inheritdoc cref="IChart.Title"/>
        /// </summary>
        public string Title
        {
            get
            {
                if (_title == null)
                {
                    _title = TryGetTitle();
                }

                return _title ?? throw new SlideDotNetException(ExceptionMessages.NotTitle);
            }
        }

        /// <summary>
        /// <inheritdoc cref="IChart.HasTitle"/>
        /// </summary>
        public bool HasTitle
        {
            get
            {
                if (_title == null)
                {
                    _title = TryGetTitle();
                }

                return _title != null;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ChartEx"/> class.
        /// </summary>
        public ChartEx(P.GraphicFrame grFrame, SlidePart sldPart)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            _sldPart = sldPart;
            _grFrame = grFrame;

            Init(); //TODO: convert to lazy loading
        }

        #endregion

        #region Private Methods

        private void Init()
        {
            // Get reference
            var chartRef = _grFrame.Descendants<C.ChartReference>().Single();

            // Get chart part by reference
            var chPart = (ChartPart)_sldPart.GetPartById(chartRef.Id);

            _cChart = chPart.ChartSpace.GetFirstChild<C.Chart>();
            _chartElements = _cChart.PlotArea.Elements().Where(e => e.LocalName.EndsWith("Chart")).ToList();
        }

        private void ParseType()
        {
            if (_chartElements.Count > 1)
            {
                _type = ChartType.Combination;
            }
            else
            {
                var chartName = _chartElements.Single().LocalName;
                Enum.TryParse(chartName, true, out ChartType chartType);
                _type = chartType;
            }
        }

        private string TryGetTitle()
        {
            var title = _cChart.Title;
            if (title == null) // chart has not title
            {
                return null;
            }
           
            var xmlChartText = title.ChartText;
            var existStatic = TryGetStatic(xmlChartText, out var staticTitle);
            if (existStatic)
            {
                return staticTitle;
            }

            // Dynamic title
            if (xmlChartText != null)
            {
                return xmlChartText.Descendants<C.StringPoint>().Single().InnerText;
            }

            if (Type == ChartType.PieChart)
            {
                // Parses PieChart dynamic title
                return _chartElements.Single().GetFirstChild<C.PieChartSeries>().GetFirstChild<C.SeriesText>().Descendants<C.StringPoint>().Single().InnerText;
            }

            return null;
        }

        #endregion

        private bool TryGetStatic(C.ChartText chartText, out string staticTitle)
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
    }
}


