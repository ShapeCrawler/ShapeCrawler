using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

internal sealed class ChartPoints : IReadOnlyList<IChartPoint>
{
    private readonly List<ChartPoint> chartPoints;

    internal ChartPoints(ChartPart chartPart, OpenXmlElement cSerXmlElement)
    {
        var numberReference = GetNumberReference(cSerXmlElement);
        var numberLiteral = GetNumberLiteral(cSerXmlElement);
        this.chartPoints = new ChartPointData(chartPart).Create(numberReference, numberLiteral);
    }

    public int Count => this.chartPoints.Count;

    public IChartPoint this[int index] => this.chartPoints[index];

    public IEnumerator<IChartPoint> GetEnumerator() => this.chartPoints.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private static NumberReference? GetNumberReference(OpenXmlElement cSerXmlElement)
    {
        var cVal = cSerXmlElement.GetFirstChild<Values>();
        if (cVal != null)
        {
            return cVal.NumberReference;
        }

        var cYVal = cSerXmlElement.GetFirstChild<YValues>();
        return cYVal?.NumberReference;
    }

    private static NumberLiteral? GetNumberLiteral(OpenXmlElement cSerXmlElement)
    {
        var cVal = cSerXmlElement.GetFirstChild<Values>();
        if (cVal != null)
        {
            return cVal.NumberLiteral;
        }

        var cYVal = cSerXmlElement.GetFirstChild<YValues>();
        return cYVal?.NumberLiteral;
    }
}