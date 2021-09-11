using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts
{
    public interface IChartPointCollection : IReadOnlyList<IChartPoint>
    {
        
    }

    internal class ChartPointCollection : IChartPointCollection
    {
        private readonly ChartPoint[] chartPoints;

        public ChartPointCollection(List<ChartPoint> points)
        {
            this.chartPoints = points.ToArray();
        }

        public IChartPoint this[int index] => this.chartPoints[index];

        public int Count => throw new System.NotImplementedException();

        public static ChartPointCollection Create(SCChart chart, OpenXmlElement cSerXmlElement)
        {
            C.NumberReference numReference;
            C.Values? cVal = cSerXmlElement.GetFirstChild<C.Values>();
            if (cVal != null)
            {
                // Some charts do have <c:val> element, for example, scatter chart.
                numReference = cVal.NumberReference!;
            }
            else
            {
                numReference = cSerXmlElement.GetFirstChild<C.YValues>() !.NumberReference!;
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
            throw new System.NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }
    }

    internal class ChartPoint : IChartPoint
    {
        public ChartPoint(double value)
        {
            this.Value = value;
        }

        public double Value { get; }
    }

    public interface IChartPoint
    {
        public double Value { get; }
    }
}