using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents collection of chart points.
    /// </summary>
    public interface IChartPointCollection : IReadOnlyList<IChartPoint>
    {
    }

    internal class ChartPointCollection : IChartPointCollection
    {
        private readonly List<ChartPoint> chartPoints;

        private ChartPointCollection(List<ChartPoint> points)
        {
            this.chartPoints = points;
        }

        public int Count => this.chartPoints.Count;

        public IChartPoint this[int index] => this.chartPoints[index];

        public IEnumerator<IChartPoint> GetEnumerator()
        {
            return this.chartPoints.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        internal static ChartPointCollection Create(SCChart chart, OpenXmlElement cSerXmlElement)
        {
            var cVal = cSerXmlElement.GetFirstChild<Values>();
            var cNumberReference = cVal != null ? cVal.NumberReference! : cSerXmlElement.GetFirstChild<YValues>() !.NumberReference!;

            // Get addresses
            var cFormula = cNumberReference.Formula!;
            var normalizedFormula = cFormula.Text.Replace("'", string.Empty).Replace("$", string.Empty);
            var dataSheetName = Regex.Match(normalizedFormula, @".+(?=\!)").Value; // eg: Sheet1!A2:A5 -> Sheet1
            var cellsRange = Regex.Match(normalizedFormula, @"(?<=\!).+").Value; // eg: Sheet1!A2:A5 -> A2:A5
            var pointAddresses = new CellsRangeParser(cellsRange).GetCellAddresses();

            // Get cached values
            List<double>? pointCachedValues = null;
            if (cNumberReference.NumberingCache != null)
            {
                var cNumericValues = cNumberReference.NumberingCache.Descendants<C.NumericValue>();
                pointCachedValues = new List<double>(cNumericValues.Count());
                foreach (var numericValue in cNumericValues)
                {
                    var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    var roundNumber = Math.Round(number, 1);
                    pointCachedValues.Add(roundNumber);
                }
            }

            // Generate points
            var chartPoints = pointAddresses.Select((address, index) => pointCachedValues == null
                    ? new ChartPoint(chart, dataSheetName, address)
                    : new ChartPoint(chart, dataSheetName, address, pointCachedValues[index]))
                .ToList();

            return new ChartPointCollection(chartPoints);
        }
    }
}