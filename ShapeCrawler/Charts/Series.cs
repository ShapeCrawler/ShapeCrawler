using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
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
        private readonly Lazy<string> _name;
        private readonly Lazy<IReadOnlyList<double>> _pointValues;
        private readonly OpenXmlElement _seriesXmlElement;

        #region Constructors

        internal Series(
            SlideChart slideChart,
            ChartType type,
            OpenXmlElement seriesXmlElement)
        {
            SlideChart = slideChart;
            _seriesXmlElement = seriesXmlElement;
            _pointValues = new Lazy<IReadOnlyList<double>>(GetPointValues);
            _name = new Lazy<string>(GetNameOrDefault);
            Type = type;
        }

        #endregion Constructors

        internal SlideChart SlideChart { get; }

        /// <summary>
        ///     Gets chart type.
        /// </summary>
        public ChartType Type { get; }

        /// <summary>
        ///     Gets collection of point values.
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


        #region Private Methods

        private IReadOnlyList<double> GetPointValues()
        {
            C.NumberReference numReference;
            C.Values cVal = _seriesXmlElement.GetFirstChild<C.Values>();
            if (cVal != null) // scatter type chart does not have <c:val> element
            {
                numReference = cVal.NumberReference;
            }
            else
            {
                numReference = _seriesXmlElement.GetFirstChild<C.YValues>().NumberReference;
            }

            return ChartReferencesParser.GetNumbersFromCacheOrSpreadsheet(numReference, SlideChart);
        }

        private string GetNameOrDefault()
        {
            C.StringReference cStringReference = _seriesXmlElement.GetFirstChild<C.SeriesText>()?.StringReference;
            if (cStringReference == null)
            {
                return null;
            }

            return ChartReferencesParser.GetSingleString(cStringReference, SlideChart);
        }

        #endregion Private Methods
    }
}