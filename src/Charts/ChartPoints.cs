using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartPoints : IReadOnlyList<IChartPoint>
{
    private readonly ChartPart chartPart;
    private readonly List<ChartPoint> chartPoints;

    internal ChartPoints(ChartPart chartPart, OpenXmlElement cSerXmlElement)
    {
        this.chartPart = chartPart;
        
        var numberReference = GetNumberReference(cSerXmlElement);
        var (sheetName, addresses) = ParseFormulaAddresses(numberReference.Formula!);
        var numericValues = GetNumericValues(numberReference);
        
        this.chartPoints = this.CreateChartPoints(addresses, numericValues, sheetName);
    }
    
    public int Count => this.chartPoints.Count;

    public IChartPoint this[int index] => this.chartPoints[index];

    public IEnumerator<IChartPoint> GetEnumerator() => this.chartPoints.GetEnumerator();
    
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
    
    private static NumberReference GetNumberReference(OpenXmlElement cSerXmlElement)
    {
        var cVal = cSerXmlElement.GetFirstChild<Values>();
        return cVal != null 
            ? cVal.NumberReference! 
            : cSerXmlElement.GetFirstChild<YValues>() !.NumberReference!;
    }

    private static (string SheetName, List<string> Addresses) ParseFormulaAddresses(Formula formula)
    {
        var normalizedFormula = formula.Text.Replace("$", string.Empty).Replace("'", string.Empty);
        var sheetName = ExtractSheetName(normalizedFormula);
        var addresses = ExtractAddresses(normalizedFormula);
        
        return (sheetName, addresses);
    }

    private static string ExtractSheetName(string normalizedFormula)
    {
        return Regex.Match(
            normalizedFormula, 
            @"(?<=\(*)[\p{L} 0-9]+?(?=!)", 
            RegexOptions.None, 
            TimeSpan.FromMilliseconds(1000)).Value; // eg: Sheet1!A2:A5 -> Sheet1
    }

    private static List<string> ExtractAddresses(string normalizedFormula)
    {
        var addressMatches = Regex.Matches(
            normalizedFormula, 
            @"[A-Z]\d+(:[A-Z]\d+)*", 
            RegexOptions.None, 
            TimeSpan.FromMilliseconds(1000)); // eg: Sheet1!A2:A5 -> A2:A5
        
        var addresses = new List<string>();
        foreach (Match match in addressMatches)
        {
            if (match.Value.Contains(":"))
            {
                var rangePointAddresses = new CellsRange(match.Value).Addresses();
                addresses.AddRange(rangePointAddresses);
            }
            else
            {
                addresses.Add(match.Value);
            }
        }
        
        return addresses;
    }

    private static List<C.NumericValue>? GetNumericValues(NumberReference numberReference)
    {
        if (numberReference.NumberingCache != null)
        {
            return [.. numberReference.NumberingCache.Descendants<C.NumericValue>()];
        }
        
        return null;
    }

    private List<ChartPoint> CreateChartPoints(List<string> addresses, List<C.NumericValue>? numericValues, string sheetName)
    {
        var points = new List<ChartPoint>(addresses.Count);

        if (addresses.Count == 1 && numericValues?.Count > 1)
        {
            foreach (var numericValue in numericValues)
            {
                points.Add(new ChartPoint(this.chartPart, numericValue, sheetName, addresses[0]));
            }
        }
        else
        {
            // Empty cells of range don't have the corresponding C.NumericValue.
            var quPoints = System.Math.Min(addresses.Count, numericValues?.Count ?? 0);
            for (int i = 0; i < quPoints; i++)
            {
                points.Add(new ChartPoint(this.chartPart, numericValues?[i]!, sheetName, addresses[i]));
            }
        }

        return points;
    }
}