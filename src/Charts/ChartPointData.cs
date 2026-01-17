using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartPointData(ChartPart chartPart)
{
    internal List<ChartPoint> Create(NumberReference? numberReference, NumberLiteral? numberLiteral)
    {
        if (numberReference?.Formula != null)
        {
            var (sheetName, addresses) = ParseFormulaAddresses(numberReference.Formula);
            var numericValues = GetNumericValues(numberReference);
            return this.CreateChartPoints(addresses, numericValues, sheetName);
        }

        return numberLiteral != null ? CreateChartPointsFromLiteral(numberLiteral) : [];
    }

    private static List<ChartPoint> CreateChartPointsFromLiteral(NumberLiteral numberLiteral)
    {
        return
        [
            .. numberLiteral
                .Elements<NumericPoint>()
                .Select(numericPoint => numericPoint.NumericValue)
                .OfType<NumericValue>()
                .Select(numericValue => new ChartPoint(numericValue))
        ];
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
                points.Add(new ChartPoint(chartPart, numericValue, sheetName, addresses[0]));
            }
        }
        else
        {
            // Empty cells of range don't have the corresponding C.NumericValue.
            var quPoints = Math.Min(addresses.Count, numericValues?.Count ?? 0);
            for (int i = 0; i < quPoints; i++)
            {
                points.Add(new ChartPoint(chartPart, numericValues?[i]!, sheetName, addresses[i]));
            }
        }

        return points;
    }
}