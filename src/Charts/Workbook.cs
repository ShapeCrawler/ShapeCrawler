using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed partial class Workbook(EmbeddedPackagePart embeddedPackagePart)
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

    /// <summary>
    ///     Gets the one-based worksheet column index from a chart formula.
    /// </summary>
    /// <param name="formula">A worksheet formula.</param>
    /// <returns>The one-based worksheet column index.</returns>
    internal int FormulaColumnIndex(string formula) => ParseFormulaAddressRange(formula).StartColumn;

    /// <summary>
    ///     Replaces the worksheet column in a chart formula while preserving its row range and sheet.
    /// </summary>
    /// <param name="formula">The source worksheet formula.</param>
    /// <param name="column">The one-based destination worksheet column.</param>
    /// <returns>The updated worksheet formula.</returns>
    internal string FormulaWithColumn(string formula, int column)
    {
        var separatorIndex = formula.LastIndexOf('!');
        var addressRange = ParseFormulaAddressRange(formula);
        var newRange = addressRange.IsSingleCell
            ? $"${ColumnName(column)}${addressRange.StartRow}"
            : $"${ColumnName(column)}${addressRange.StartRow}:${ColumnName(column)}${addressRange.EndRow}";
        return $"{formula[..(separatorIndex + 1)]}{newRange}";
    }

    /// <summary>
    ///     Appends a series name and values to the embedded workbook.
    /// </summary>
    /// <param name="nameFormula">The worksheet cell formula for the series name.</param>
    /// <param name="valuesFormula">The worksheet range formula for series values.</param>
    /// <param name="name">The name to write to the worksheet.</param>
    /// <param name="values">The values to write to the worksheet.</param>
    internal void AppendSeriesColumn(string nameFormula, string valuesFormula, string name, IReadOnlyList<double> values)
    {
        var (nameSheet, _) = ParseFormula(nameFormula);
        var (valuesSheet, _) = ParseFormula(valuesFormula);
        var nameRange = ParseFormulaAddressRange(nameFormula);
        var valuesRange = ParseFormulaAddressRange(valuesFormula);
        if (!string.Equals(nameSheet, valuesSheet, StringComparison.OrdinalIgnoreCase)
            || !nameRange.IsSingleCell
            || valuesRange.StartColumn != valuesRange.EndColumn
            || valuesRange.EndRow - valuesRange.StartRow + 1 != values.Count
            || nameRange.StartColumn != valuesRange.StartColumn)
        {
            throw new SCException("The template series must use a single header cell and a contiguous vertical values range on one worksheet.");
        }

        using var transactionStream = new MemoryStream();
        using (var sourceStream = embeddedPackagePart.GetStream(FileMode.Open, FileAccess.Read))
        {
            sourceStream.CopyTo(transactionStream);
        }

        transactionStream.Position = 0;
        using (var spreadsheetDocument = SpreadsheetDocument.Open(transactionStream, true))
        {
            var workbookPart = spreadsheetDocument.WorkbookPart!;
            var sheet = workbookPart.Workbook!.Sheets!.Elements<X.Sheet>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, nameSheet, StringComparison.OrdinalIgnoreCase))
                ?? throw new SCException($"Could not find worksheet '{nameSheet}'.");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var worksheet = worksheetPart.Worksheet!;
            var sheetData = worksheet.GetFirstChild<X.SheetData>()!;
            var targetColumn = nameRange.StartColumn;
            var headerAddress = $"{ColumnName(targetColumn)}{nameRange.StartRow}";
            var targetAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { headerAddress };
            for (var index = 0; index < values.Count; index++)
            {
                targetAddresses.Add($"{ColumnName(targetColumn)}{valuesRange.StartRow + index}");
            }

            if (worksheet.Descendants<X.Cell>().Any(cell => targetAddresses.Contains(cell.CellReference?.Value ?? string.Empty)))
            {
                throw new SCException($"The target worksheet column '{ColumnName(targetColumn)}' already contains chart data.");
            }

            AddCell(sheetData, headerAddress, name, true);
            for (var index = 0; index < values.Count; index++)
            {
                var address = $"{ColumnName(targetColumn)}{valuesRange.StartRow + index}";
                AddCell(sheetData, address, values[index].ToString(CultureInfo.InvariantCulture));
            }

            UpdateWorksheetDimension(worksheet, targetColumn, valuesRange.EndRow);
            UpdateTableDefinitions(worksheetPart, targetColumn, name);
            worksheet.Save();
        }

        using var updatedStream = new MemoryStream(transactionStream.ToArray());
        embeddedPackagePart.FeedData(updatedStream);
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

    private static AddressRange ParseFormulaAddressRange(string formula)
    {
        var (_, range) = ParseFormula(formula);
        return ParseAddressRange(range);
    }

    private static AddressRange ParseAddressRange(string range)
    {
        var addresses = range.Split(':');
        if (addresses.Length is < 1 or > 2)
        {
            throw new SCException($"Unsupported worksheet range: {range}.");
        }

        var start = ParseAddress(addresses[0]);
        var end = addresses.Length == 1 ? start : ParseAddress(addresses[1]);
        if (start.Column > end.Column || start.Row > end.Row)
        {
            throw new SCException($"Unsupported worksheet range: {range}.");
        }

        return new AddressRange(start.Column, start.Row, end.Column, end.Row);
    }

    private static (int Column, int Row) ParseAddress(string address)
    {
        var normalized = address.Replace("$", string.Empty);
        var match = System.Text.RegularExpressions.Regex.Match(
            normalized,
            "^([A-Z]+)([0-9]+)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase,
            TimeSpan.FromSeconds(1));
        if (!match.Success
            || !int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var row))
        {
            throw new SCException($"Unsupported worksheet cell address: {address}.");
        }

        var column = 0;
        foreach (var character in match.Groups[1].Value.ToUpperInvariant())
        {
            column = (column * 26) + character - 'A' + 1;
        }

        return (column, row);
    }

    private static string ColumnName(int column)
    {
        var result = string.Empty;
        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            result = (char)('A' + remainder) + result;
            column = (column - 1) / 26;
        }

        return result;
    }

    private static void AddCell(X.SheetData sheetData, string address, string value, bool isText = false)
    {
        var (column, rowNumber) = ParseAddress(address);
        var row = sheetData.Elements<X.Row>().FirstOrDefault(item => item.RowIndex?.Value == (uint)rowNumber);
        if (row == null)
        {
            row = new X.Row { RowIndex = (uint)rowNumber };
            var followingRow = sheetData.Elements<X.Row>().FirstOrDefault(item => item.RowIndex?.Value > (uint)rowNumber);
            sheetData.InsertBefore(row, followingRow);
        }

        var cell = isText
            ? new X.Cell
            {
                CellReference = address,
                DataType = X.CellValues.InlineString,
                InlineString = new X.InlineString(new X.Text(value))
            }
            : new X.Cell
            {
                CellReference = address,
                CellValue = new X.CellValue(value)
            };
        var followingCell = row.Elements<X.Cell>().FirstOrDefault(item => ParseAddress(item.CellReference!.Value!).Column > column);
        row.InsertBefore(cell, followingCell);
    }

    private static void UpdateWorksheetDimension(X.Worksheet worksheet, int column, int row)
    {
        var dimension = worksheet.SheetDimension;
        if (dimension?.Reference?.Value == null)
        {
            return;
        }

        var current = ParseAddressRange(dimension.Reference.Value);
        var endColumn = Math.Max(column, current.EndColumn);
        var endRow = Math.Max(row, current.EndRow);
        dimension.Reference = $"{ColumnName(current.StartColumn)}{current.StartRow}:{ColumnName(endColumn)}{endRow}";
    }

    private static void UpdateTableDefinitions(WorksheetPart worksheetPart, int column, string seriesName)
    {
        foreach (var tablePart in worksheetPart.TableDefinitionParts)
        {
            var table = tablePart.Table;
            if (table?.Reference?.Value == null)
            {
                continue;
            }

            var range = ParseAddressRange(table.Reference.Value);
            if (range.EndColumn != column - 1)
            {
                continue;
            }

            var columns = table.GetFirstChild<X.TableColumns>();
            if (columns == null)
            {
                continue;
            }

            if (columns.Elements<X.TableColumn>().Any(item =>
                    string.Equals(item.Name?.Value, seriesName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new SCException($"The embedded workbook already contains a column named '{seriesName}'.");
            }

            var columnId = columns.Elements<X.TableColumn>().Select(item => item.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;
            columns.AppendChild(new X.TableColumn { Id = columnId, Name = seriesName });
            columns.Count = (uint)columns.Elements<X.TableColumn>().Count();
            table.Reference = $"{ColumnName(range.StartColumn)}{range.StartRow}:{ColumnName(column)}{range.EndRow}";
            table.AutoFilter?.Reference = table.Reference.Value;

            table.Save();
        }
    }

    private readonly record struct AddressRange(int StartColumn, int StartRow, int EndColumn, int EndRow)
    {
        /// <summary>
        ///     Gets a value indicating whether the range contains one cell.
        /// </summary>
        internal bool IsSingleCell => this.StartColumn == this.EndColumn && this.StartRow == this.EndRow;
    }
}