using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

/// <summary>
///     Provides access to the chart's embedded workbook.
/// </summary>
internal sealed partial class Workbook
{
    /// <summary>
    ///     Extends a vertical chart data formula by one point.
    /// </summary>
    /// <param name="formula">The existing worksheet range formula.</param>
    /// <param name="count">The existing chart point count.</param>
    /// <returns>The extended worksheet range formula.</returns>
    internal string ExtendedPointFormula(string formula, uint count)
    {
        var range = ParseFormulaAddressRange(formula);
        if (range.StartColumn != range.EndColumn || range.EndRow - range.StartRow + 1 != count)
        {
            throw new SCException("Chart data must use contiguous vertical worksheet ranges matching the chart point count.");
        }

        var prefix = formula[..(formula.LastIndexOf('!') + 1)];
        var column = ColumnName(range.StartColumn);
        return $"{prefix}${column}${range.StartRow}:${column}${range.EndRow + 1}";
    }

    /// <summary>
    ///     Writes chart points after their existing worksheet ranges.
    /// </summary>
    /// <param name="cells">The original range formulas and their new cell values.</param>
    internal void AppendPointCells(IReadOnlyList<(string Formula, string Value, bool IsText)> cells)
    {
        if (cells.Count == 0)
        {
            return;
        }

        using var transactionStream = new MemoryStream();
        using (var sourceStream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read))
        {
            sourceStream.CopyTo(transactionStream);
        }

        transactionStream.Position = 0;
        using (var document = SpreadsheetDocument.Open(transactionStream, true))
        {
            var writtenCells = new Dictionary<string, (string Value, bool IsText)>(StringComparer.OrdinalIgnoreCase);
            foreach (var (formula, value, isText) in cells)
            {
                AppendPointCell(document.WorkbookPart!, formula, value, isText, writtenCells);
            }
        }

        using var updatedStream = new MemoryStream(transactionStream.ToArray());
        embeddedPackagePart.FeedData(updatedStream);
    }

    private static void AppendPointCell(
        WorkbookPart workbookPart,
        string formula,
        string value,
        bool isText,
        Dictionary<string, (string Value, bool IsText)> writtenCells)
    {
        var (sheetName, _) = ParseFormula(formula);
        var range = ParseFormulaAddressRange(formula);
        var address = $"{ColumnName(range.EndColumn)}{range.EndRow + 1}";
        var key = $"{sheetName}!{address}";
        if (writtenCells.TryGetValue(key, out var previous))
        {
            if (previous != (value, isText))
            {
                throw new SCException($"Conflicting chart values target worksheet cell '{key}'.");
            }

            return;
        }

        var sheet = workbookPart.Workbook!.Sheets!.Elements<X.Sheet>()
            .FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new SCException($"Could not find worksheet '{sheetName}'.");
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var worksheet = worksheetPart.Worksheet!;
        if (worksheet.Descendants<X.Cell>().Any(cell => string.Equals(cell.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase)))
        {
            throw new SCException($"The target worksheet cell '{key}' already exists.");
        }

        AddCell(worksheet.GetFirstChild<X.SheetData>()!, address, value, isText);
        UpdateWorksheetDimension(worksheet, range.EndColumn, range.EndRow + 1);
        ExtendPointTables(worksheetPart, range);
        worksheet.Save();
        writtenCells.Add(key, (value, isText));
    }

    private static void ExtendPointTables(WorksheetPart worksheetPart, AddressRange dataRange)
    {
        foreach (var tablePart in worksheetPart.TableDefinitionParts)
        {
            var table = tablePart.Table;
            if (table?.Reference?.Value == null)
            {
                continue;
            }

            var range = ParseAddressRange(table.Reference.Value);
            if (range.EndRow != dataRange.EndRow
                || dataRange.StartColumn < range.StartColumn
                || dataRange.EndColumn > range.EndColumn)
            {
                continue;
            }

            if (table.TotalsRowCount?.Value > 0)
            {
                throw new SCException("Appending chart data to a worksheet table with a totals row is not supported.");
            }

            table.Reference = $"{ColumnName(range.StartColumn)}{range.StartRow}:{ColumnName(range.EndColumn)}{range.EndRow + 1}";
            table.AutoFilter?.Reference = table.Reference.Value;
            table.Save();
        }
    }
}