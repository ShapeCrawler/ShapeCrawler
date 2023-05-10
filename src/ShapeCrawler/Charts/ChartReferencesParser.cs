using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Charts;

internal static class ChartReferencesParser
{
    internal static IEnumerable<double> GetNumbersFromCacheOrWorkbook(C.NumberReference numberReference, SCChart slideChart)
    {
        if (numberReference.NumberingCache != null)
        {
            // From cache
            var cNumericValues = numberReference.NumberingCache.Descendants<C.NumericValue>();
            var cachedPointValues = new List<double>(cNumericValues.Count());
            foreach (var numericValue in cNumericValues)
            {
                var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                var roundNumber = Math.Round(number, 1);
                cachedPointValues.Add(roundNumber);
            }

            return cachedPointValues;
        }

        // From Spreadsheet
        var rangeXCells = GetXCellsByFormula(numberReference.Formula!, slideChart);
        var pointValues = new List<double>(rangeXCells.Count);
        foreach (var xCell in rangeXCells)
        {
            var cellValue = xCell.InnerText.Length == 0 ? 0 : double.Parse(xCell.InnerText, CultureInfo.InvariantCulture.NumberFormat);
            pointValues.Add(cellValue);
        }

        return pointValues;
    }

    internal static string GetSingleString(C.StringReference stringReference, SCChart slideChart)
    {
        string fromCache = stringReference.StringCache?.GetFirstChild<C.StringPoint>() !.Single().InnerText!;
        if (fromCache != null)
        {
            return fromCache;
        }

        List<X.Cell> xCell = GetXCellsByFormula(stringReference.Formula!, slideChart);

        return xCell.Single().InnerText;
    }

    /// <summary>
    ///     Gets cell values.
    /// </summary>
    internal static List<X.Cell> GetXCellsByFormula(C.Formula cFormula, SCChart chart)
    {
        var normalizedFormula = cFormula.Text.Replace("'", string.Empty).Replace("$", string.Empty); // eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5
        var chartSheetName = Regex.Match(normalizedFormula, @".+(?=\!)").Value; // eg: Sheet1!A2:A5 -> Sheet1
        var cellsRange = Regex.Match(normalizedFormula, @"(?<=\!).+").Value; // eg: Sheet1!A2:A5 -> A2:A5

        var workbookPart = chart.ChartWorkbook!.WorkbookPart;
        var chartSheet = workbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == chartSheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(chartSheet.Id!);
        var sheetXCells = worksheetPart.Worksheet.Descendants<X.Cell>();

        var rangeCellAddresses = new CellsRangeParser(cellsRange).GetCellAddresses();
        var rangeXCells = new List<X.Cell>(rangeCellAddresses.Count);
        foreach (var address in rangeCellAddresses)
        {
            var xCell = sheetXCells.First(xCell => xCell.CellReference == address);
            rangeXCells.Add(xCell);
        }

        return rangeXCells;
    }
}