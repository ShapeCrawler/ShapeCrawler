using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class SeriesXPoints : IReadOnlyList<IChartPoint>
{
    private readonly List<ChartPoint> chartPoints;

    internal SeriesXPoints(ChartPart chartPart, OpenXmlElement cSerXmlElement)
    {
        var cXValues = cSerXmlElement.GetFirstChild<C.XValues>();
        if (cXValues == null)
        {
            this.chartPoints = [];
            return;
        }

        var numberReference = cXValues.NumberReference;
        var numberLiteral = cXValues.NumberLiteral;
        this.chartPoints = new ChartPointData(chartPart).Create(numberReference, numberLiteral);
    }

    public int Count => this.chartPoints.Count;

    public IChartPoint this[int index] => this.chartPoints[index];

    public IEnumerator<IChartPoint> GetEnumerator() => this.chartPoints.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
}