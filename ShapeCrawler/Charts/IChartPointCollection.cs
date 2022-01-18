using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

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

        internal static ChartPointCollection Create(SCChart chart, OpenXmlElement cSerXmlElement)
        {
            var cVal = cSerXmlElement.GetFirstChild<Values>();
            var cNumberReference = cVal != null ? cVal.NumberReference! : cSerXmlElement.GetFirstChild<YValues>() !.NumberReference!;

            var pointValues = ChartReferencesParser.GetNumbersFromCacheOrWorkbook(cNumberReference, chart);
            var chartPoints = pointValues.Select(point => new ChartPoint(point)).ToList();

            return new ChartPointCollection(chartPoints);
        }

        public IEnumerator<IChartPoint> GetEnumerator()
        {
            return this.chartPoints.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}