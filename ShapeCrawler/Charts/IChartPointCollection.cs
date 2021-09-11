using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Spreadsheet;

namespace ShapeCrawler.Charts
{
    public interface IChartPointCollection : IReadOnlyList<IChartPoint>
    {
        
    }

    internal class ChartPointCollection : IChartPointCollection
    {
        private readonly List<ChartPoint> chartPoints;

        public ChartPointCollection(List<ChartPoint> points)
        {
            this.chartPoints = points;
        }

        public IChartPoint this[int index] => this.chartPoints[index];

        public int Count => throw new System.NotImplementedException();

        public static ChartPointCollection Create(SCChart chart, OpenXmlElement cSerXmlElement)
        {
            DocumentFormat.OpenXml.Drawing.Charts.NumberReference numReference;
            DocumentFormat.OpenXml.Drawing.Charts.Values? cVal = cSerXmlElement.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Values>();
            if (cVal != null)
            {
                // Some charts do have <c:val> element, for example, scatter chart.
                numReference = cVal.NumberReference!;
            }
            else
            {
                numReference = cSerXmlElement.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.YValues>()!.NumberReference!;
            }

            IReadOnlyList<double> pointValues = ChartReferencesParser.GetNumbersFromCacheOrWorkbook(numReference, chart);
            List<ChartPoint> points = new();
            foreach (double point in pointValues)
            {
                points.Add(new ChartPoint(point));
            }

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