using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using ShapeCrawler.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Charts
{
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

        public IChartPoint this[int index] => this.chartPoints[index];

        public int Count => throw new System.NotImplementedException();

        public static ChartPointCollection Create(SCChart chart, OpenXmlElement cSerXmlElement)
        {
            var cVal = cSerXmlElement.GetFirstChild<Values>();
            var numReference = cVal != null ? cVal.NumberReference! : cSerXmlElement.GetFirstChild<YValues>()!.NumberReference!;

            IReadOnlyList<double> pointValues = ChartReferencesParser.GetNumbersFromCacheOrWorkbook(numReference, chart);
            List<ChartPoint> points = pointValues.Select(point => new ChartPoint(point)).ToList();

            return new ChartPointCollection(points);
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