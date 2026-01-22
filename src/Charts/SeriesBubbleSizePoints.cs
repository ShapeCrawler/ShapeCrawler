using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

/// <summary>
///     Represents bubble size points of a series.
/// </summary>
internal sealed class SeriesBubbleSizePoints : IReadOnlyList<IChartPoint>
{
    private readonly List<ChartPoint> chartPoints;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SeriesBubbleSizePoints"/> class.
    /// </summary>
    /// <param name="chartPart">The chart part owning the series.</param>
    /// <param name="cSerXmlElement">The series Open XML element.</param>
    internal SeriesBubbleSizePoints(ChartPart chartPart, OpenXmlElement cSerXmlElement)
    {
        var cBubbleSize = cSerXmlElement.GetFirstChild<C.BubbleSize>();
        if (cBubbleSize == null)
        {
            this.chartPoints = [];
            return;
        }

        var numberReference = cBubbleSize.NumberReference;
        var numberLiteral = cBubbleSize.NumberLiteral;
        this.chartPoints = new ChartPointData(chartPart).Create(numberReference, numberLiteral);
    }

    /// <inheritdoc />
    public int Count => this.chartPoints.Count;

    /// <inheritdoc />
    public IChartPoint this[int index] => this.chartPoints[index];

    /// <inheritdoc />
    public IEnumerator<IChartPoint> GetEnumerator() => this.chartPoints.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
}