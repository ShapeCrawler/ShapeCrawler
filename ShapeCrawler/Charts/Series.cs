using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Enums;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using ShapeCrawler.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Charts
{
    /// <summary>
    /// Represents a chart series.
    /// </summary>
    public class Series
    {
        #region Fields

        private readonly Lazy<List<double>> _pointValues;
        private readonly Lazy<string> _name;
        private readonly ChartPart _sdkChartPart;
        private readonly OpenXmlElement _sdkSeries;
        private readonly ChartRefParser _chartRefParser;

        #endregion Fields

        /// <summary>
        /// Returns a chart type.
        /// </summary>
        public ChartType Type { get; }

        /// <summary>
        /// Returns a point values.
        /// </summary>
        public IList<double> PointValues => _pointValues.Value;

        public bool HasName => _name.Value != null;

        public string Name
        {
            get
            {
                if (_name.Value == null)
                {
                    throw new NotSupportedException(ExceptionMessages.SeriesHasNotName);
                }

                return _name.Value;
            }
        }

        #region Constructors

        public Series(ChartType type, OpenXmlElement sdkSeries, ChartPart sdkChartPart, ChartRefParser chartRefParser)
        {
            _sdkSeries = sdkSeries ?? throw new ArgumentNullException(nameof(sdkSeries));
            Check.NotNull(sdkSeries, nameof(sdkSeries));
            Check.NotNull(sdkChartPart, nameof(sdkChartPart));

            _sdkChartPart = sdkChartPart ?? throw new ArgumentNullException(nameof(sdkChartPart));
            _chartRefParser = chartRefParser;
            _pointValues = new Lazy<List<double>>(GetPointValues);
            _name = new Lazy<string>(GetNameOrDefault);
            Type = type;
        }

        #endregion Constructors

        #region Private Methods

        private List<double> GetPointValues()
        {
            C.NumberReference numReference;
            var cVal = _sdkSeries.GetFirstChild<C.Values>();
            if (cVal != null) // scatter type chart does not have <c:val> element
            {
                numReference = cVal.NumberReference;
            }
            else
            {
                numReference = _sdkSeries.GetFirstChild<C.YValues>().NumberReference;
            }

            return _chartRefParser.GetNumbers(numReference, _sdkChartPart).ToList(); //TODO: remove to list
        }

        private string GetNameOrDefault()
        {
            var strReference = _sdkSeries.GetFirstChild<C.SeriesText>()?.StringReference;
            if (strReference == null)
            {
                return null;
            }

            return _chartRefParser.GetSingleString(strReference, _sdkChartPart);
        }

        #endregion Private Methods
    }
}