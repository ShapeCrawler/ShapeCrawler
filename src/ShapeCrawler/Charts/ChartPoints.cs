using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

internal sealed class ChartPoints : IReadOnlyList<IChartPoint>
{
    private readonly ChartPart sdkChartPart;
    private readonly List<ChartPoint> chartPoints;

    internal ChartPoints(ChartPart sdkChartPart, OpenXmlElement cSerXmlElement)
    {
        this.sdkChartPart = sdkChartPart;
        
        var cVal = cSerXmlElement.GetFirstChild<Values>();
        var cNumberReference =
            cVal != null ? cVal.NumberReference! : cSerXmlElement.GetFirstChild<YValues>() !.NumberReference!;

        // Get addresses
        var cFormula = cNumberReference.Formula!;
        
        var normalizedFormula = cFormula.Text.Replace("$", string.Empty).Replace("'", string.Empty);
        var sheetName =
            Regex.Match(normalizedFormula, @"(?<=\(*)[\p{L} 0-9]+?(?=!)").Value; // eg: Sheet1!A2:A5 -> Sheet1
        var addressMatches = Regex.Matches(normalizedFormula, @"[A-Z]\d+(:[A-Z]\d+)*"); // eg: Sheet1!A2:A5 -> A2:A5
        var addresses = new List<string>();
        foreach (Match match in addressMatches)
        {
            if (match.Value.Contains(':'))
            {
                var rangePointAddresses = new ExcelCellsRange(match.Value).Addresses();
                addresses.AddRange(rangePointAddresses);
            }
            else
            {
                addresses.Add(match.Value);
            }
        }

        // Get cached values
        List<C.NumericValue>? cNumericValues = null;
        if (cNumberReference.NumberingCache != null)
        {
            cNumericValues = cNumberReference.NumberingCache.Descendants<C.NumericValue>().ToList();
        }

        // Generate points
        var chartPoints = new List<ChartPoint>(addresses.Count);

        if (addresses.Count == 1 && cNumericValues?.Count > 1)
        {
            foreach (var cNumericValue in cNumericValues)
            {
                chartPoints.Add(new ChartPoint(this.sdkChartPart, cNumericValue, sheetName, addresses[0]));
            }
        }
        else
        {
            // Empty cells of range don't have the corresponding C.NumericValue.
            var quPoints = System.Math.Min(addresses.Count, cNumericValues?.Count ?? 0);
            for (int i = 0; i < quPoints; i++)
            {
                chartPoints.Add(new ChartPoint(this.sdkChartPart, cNumericValues?[i]!, sheetName, addresses[i]));
            }
        }

        this.chartPoints = chartPoints;
    }

    public int Count => this.chartPoints.Count;

    public IChartPoint this[int index] => this.chartPoints[index];

    public IEnumerator<IChartPoint> GetEnumerator() => this.chartPoints.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
}