using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Spreadsheet;
using SlideDotNet.Validation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models.SlideComponents.Chart
{
    /// <summary>
    /// Represents a chart series.
    /// </summary>
    public class Series
    {
        #region Fields

        private readonly Lazy<List<double>> _pointValues;
        private readonly ChartPart _sdkChartPart;

        #endregion Fields

        /// <summary>
        /// Returns a chart type.
        /// </summary>
        public ChartType Type { get; }

        /// <summary>
        /// Returns a point values.
        /// </summary>
        public IList<double> PointValues => _pointValues.Value; //TODO: delete setter

        #region Constructors

        public Series(ChartType type, OpenXmlElement sdkSeries, ChartPart sdkChartPart)
        {
            Check.NotNull(sdkSeries, nameof(sdkSeries));
            _sdkChartPart = sdkChartPart ?? throw new ArgumentNullException(nameof(sdkChartPart));
            _pointValues = new Lazy<List<double>>(GetPointValues(sdkSeries));
            Type = type;
        }

        #endregion Constructors

        #region Private Methods

        private List<double> GetPointValues(OpenXmlElement sdkSeries)
        {
            C.NumberReference numberReference;
            var cVal = sdkSeries.GetFirstChild<C.Values>();
            if (cVal != null) // scatter type chart does not have <c:val> element
            {
                numberReference = cVal.NumberReference;
            }
            else
            {
                numberReference = sdkSeries.GetFirstChild<C.YValues>().NumberReference;
            }

            var numberingCache = numberReference.NumberingCache;
            if (numberingCache != null)
            {
                return PointValueParser.FromCache(numberingCache).ToList(); //TODO: remove ToList()
            }

            var embeddedPackagePart = _sdkChartPart.EmbeddedPackagePart;
            var values = PointValueParser.FromFormula(numberReference.Formula, embeddedPackagePart).ToList(); //TODO: remove ToList()
            return values;
        }

        #endregion Private Methods
    }
}