using DocumentFormat.OpenXml;
using ShapeCrawler.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts
{
    public interface IChartPointCollection
    {
    }

    internal class ChartPointCollection : IChartPointCollection
    {
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

            return ChartReferencesParser.GetNumbersFromCacheOrWorkbook(numReference, chart);
        }
    }

    public interface IChartPoint
    {
    }
}