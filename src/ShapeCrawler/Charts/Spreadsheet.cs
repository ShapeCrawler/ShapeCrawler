using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class Spreadsheet(ChartPart chartPart)
{
    internal Sheet Sheet(string sheetName) => new(chartPart.EmbeddedPackagePart!, sheetName);

    internal List<double> FormulaValues(string formula)
    {
        var normalizedFormula = formula.Replace("'", string.Empty).Replace("$", string.Empty); // eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5
        var sheetName = Regex.Match(normalizedFormula, @".+(?=\!)", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value; // eg: Sheet1!A2:A5 -> Sheet1
        var cellsRange = Regex.Match(normalizedFormula, @"(?<=\!).+", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value; // eg: Sheet1!A2:A5 -> A2:A5

        var stream = chartPart.EmbeddedPackagePart!.GetStream();
        var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var sdkWorkbookPart = sdkSpreadsheetDocument.WorkbookPart!;
        var sdkSheet = sdkWorkbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == sheetName);
        var sdkWorksheetPart = (WorksheetPart)sdkWorkbookPart.GetPartById(sdkSheet.Id!);
        var sheetXCells = sdkWorksheetPart.Worksheet.Descendants<X.Cell>();

        var addresses = new CellsRange(cellsRange).Addresses();
        var rangeXCells = new List<X.Cell>(addresses.Count);
        foreach (var address in addresses)
        {
            var xCell = sheetXCells.First(xCell => xCell.CellReference == address);
            rangeXCells.Add(xCell);
        }

        var pointValues = new List<double>(rangeXCells.Count);
        foreach (var xCell in rangeXCells)
        {
            var cellValue = xCell.InnerText.Length == 0 ? 0 : double.Parse(xCell.InnerText, CultureInfo.InvariantCulture.NumberFormat);
            pointValues.Add(cellValue);
        }

        sdkSpreadsheetDocument.Dispose();
        stream.Close();

        return pointValues;
    }

    internal byte[] AsByteArray()
    {
        var stream = chartPart.EmbeddedPackagePart!.GetStream();
        var mStream = new MemoryStream();
        stream.CopyTo(mStream);

        return mStream.ToArray();
    }
}