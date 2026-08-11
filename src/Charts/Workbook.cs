using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class Workbook(EmbeddedPackagePart embeddedPackagePart)
{
    internal Worksheet Sheet(string sheetName) => new(embeddedPackagePart, sheetName);

    internal List<double> FormulaValues(string formula)
    {
        var (sheetName, cellsRange) = ParseFormula(formula);
        using var stream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var sdkWorkbookPart = sdkSpreadsheetDocument.WorkbookPart!;
        var sheetXCells = WorksheetCells(sdkWorkbookPart, sheetName);

        var addresses = new CellsRange(cellsRange).Addresses();
        var rangeXCells = new List<X.Cell>(addresses.Count);
        foreach (var address in addresses)
        {
            var xCell = WorksheetCell(sheetXCells, sheetName, address);
            rangeXCells.Add(xCell);
        }

        var pointValues = new List<double>(rangeXCells.Count);
        foreach (var xCell in rangeXCells)
        {
            var cellValue = xCell.InnerText.Length == 0
                ? 0
                : double.Parse(xCell.InnerText, CultureInfo.InvariantCulture.NumberFormat);
            pointValues.Add(cellValue);
        }

        return pointValues;
    }

    internal string FormulaText(string formula)
    {
        var (sheetName, cellsRange) = ParseFormula(formula);
        var address = SingleCellAddress(cellsRange);
        using var stream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read);
        using var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var sdkWorkbookPart = sdkSpreadsheetDocument.WorkbookPart!;
        var xCell = WorksheetCell(WorksheetCells(sdkWorkbookPart, sheetName), sheetName, address);
        var value = xCell.CellValue?.Text ?? xCell.InlineString?.InnerText ?? string.Empty;
        if (xCell.DataType?.Value != X.CellValues.SharedString)
        {
            return value;
        }

        var sharedStrings = sdkWorkbookPart.SharedStringTablePart?.SharedStringTable;
        return sharedStrings == null
            ? value
            : sharedStrings.Elements<X.SharedStringItem>()
                .ElementAt(int.Parse(value, CultureInfo.InvariantCulture))
                .InnerText;
    }

    internal void UpdateFormulaCell(string formula, string value)
    {
        var (sheetName, cellsRange) = ParseFormula(formula);
        var address = SingleCellAddress(cellsRange);
        using var transactionStream = new MemoryStream();
        using (var sourceStream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read))
        {
            sourceStream.CopyTo(transactionStream);
        }

        transactionStream.Position = 0;
        using (var sdkSpreadsheetDocument = SpreadsheetDocument.Open(transactionStream, true))
        {
            var sdkWorkbookPart = sdkSpreadsheetDocument.WorkbookPart!;
            var xCell = WorksheetCell(WorksheetCells(sdkWorkbookPart, sheetName), sheetName, address);
            xCell.RemoveAllChildren<X.InlineString>();
            xCell.CellFormula?.Remove();
            xCell.DataType = new EnumValue<X.CellValues>(X.CellValues.String);
            xCell.CellValue = new X.CellValue(value);
        }

        using var updatedStream = new MemoryStream(transactionStream.ToArray());
        embeddedPackagePart.FeedData(updatedStream);
    }

    internal byte[] AsByteArray()
    {
        using var stream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read);
        var mStream = new MemoryStream();
        stream.CopyTo(mStream);

        return mStream.ToArray();
    }

    private static (string SheetName, string CellsRange) ParseFormula(string formula)
    {
        var normalizedFormula = formula.Trim();
        if (normalizedFormula.StartsWith("=", StringComparison.Ordinal))
        {
            normalizedFormula = normalizedFormula[1..];
        }

        var separatorIndex = normalizedFormula.LastIndexOf('!');
        if (separatorIndex < 1 || separatorIndex == normalizedFormula.Length - 1)
        {
            throw new SCException($"Unsupported worksheet formula: {formula}.");
        }

        var sheetName = normalizedFormula[..separatorIndex];
        if (sheetName.Length > 1
            && sheetName.StartsWith("'", StringComparison.Ordinal)
            && sheetName.EndsWith("'", StringComparison.Ordinal))
        {
            sheetName = sheetName[1..^1].Replace("''", "'");
        }

        var workbookNameEndIndex = sheetName.LastIndexOf(']');
        if (sheetName.StartsWith("[", StringComparison.Ordinal) && workbookNameEndIndex >= 0)
        {
            sheetName = sheetName[(workbookNameEndIndex + 1)..];
        }

        var cellsRange = normalizedFormula[(separatorIndex + 1)..].Replace("$", string.Empty);
        return (sheetName, cellsRange);
    }

    private static string SingleCellAddress(string cellsRange)
    {
        var rangeSeparatorIndex = cellsRange.IndexOf(':');
        if (rangeSeparatorIndex >= 0)
        {
            var firstAddress = cellsRange[..rangeSeparatorIndex];
            var lastAddress = cellsRange[(rangeSeparatorIndex + 1)..];
            if (firstAddress.Length == 0
                || !string.Equals(firstAddress, lastAddress, StringComparison.OrdinalIgnoreCase))
            {
                throw new SCException(
                    $"Expected a single-cell worksheet formula, but found {cellsRange}.");
            }

            cellsRange = firstAddress;
        }

        if (cellsRange.IndexOf(',') >= 0)
        {
            throw new SCException(
                $"Expected a single-cell worksheet formula, but found {cellsRange}.");
        }

        var addresses = new CellsRange(cellsRange).Addresses();
        return addresses.Count == 1
            && string.Equals(addresses[0], cellsRange, StringComparison.OrdinalIgnoreCase)
            ? addresses[0]
            : throw new SCException($"Expected a single-cell worksheet formula, but found {cellsRange}.");
    }

    private static IEnumerable<X.Cell> WorksheetCells(WorkbookPart sdkWorkbookPart, string sheetName)
    {
        var sdkSheet = sdkWorkbookPart.Workbook!.Sheets!.Elements<X.Sheet>()
            .FirstOrDefault(xSheet => string.Equals(xSheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new SCException($"Could not find worksheet '{sheetName}'.");
        var sdkWorksheetPart = (WorksheetPart)sdkWorkbookPart.GetPartById(sdkSheet.Id!);
        return sdkWorksheetPart.Worksheet!.Descendants<X.Cell>();
    }

    private static X.Cell WorksheetCell(IEnumerable<X.Cell> worksheetCells, string sheetName, string address)
    {
        return worksheetCells.FirstOrDefault(cell => string.Equals(cell.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase))
            ?? throw new SCException($"Could not find cell '{address}' in worksheet '{sheetName}'.");
    }
}