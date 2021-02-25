using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart series.
    /// </summary>
    public class Series
    {
        #region Constructors

        internal Series(ChartType type, OpenXmlElement sdkSeries, ChartPart sdkChartPart, ChartReferencesParser chartRefParser)
        {
            _sdkSeries = sdkSeries ?? throw new ArgumentNullException(nameof(sdkSeries));
            _sdkChartPart = sdkChartPart ?? throw new ArgumentNullException(nameof(sdkChartPart));
            _chartRefParser = chartRefParser;
            _pointValues = new Lazy<IReadOnlyList<double>>(GetPointValues);
            _name = new Lazy<string>(GetNameOrDefault);
            Type = type;
        }

        #endregion Constructors

        /// <summary>
        ///     Returns a chart type.
        /// </summary>
        public ChartType Type { get; }

        /// <summary>
        ///     Returns a point values.
        /// </summary>
        public IReadOnlyList<double> PointValues => _pointValues.Value;

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

        #region Fields

        private readonly Lazy<IReadOnlyList<double>> _pointValues;
        private readonly Lazy<string> _name;
        private readonly ChartPart _sdkChartPart;
        private readonly OpenXmlElement _sdkSeries;
        private readonly ChartReferencesParser _chartRefParser;

        #endregion Fields

        #region Private Methods

        private IReadOnlyList<double> GetPointValues()
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

            return _chartRefParser.GetNumbersFromCacheOrSpreadsheet(numReference, _sdkChartPart);
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