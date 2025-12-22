using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class Worksheet(EmbeddedPackagePart embeddedPackagePart, string sheetName)
{
    internal WorksheetCell Cell(string address) => new(embeddedPackagePart, sheetName, address);

    internal void UpdateCell(string address, string value) => this.Cell(address).UpdateValue(value);

    internal void UpdateCell(string address, string value, X.CellValues type) => this.Cell(address).UpdateValue(value, type);
}

internal sealed class WorksheetCell(EmbeddedPackagePart embeddedPackagePart, string sheetName, string address)
{
    internal void UpdateValue(string value) => this.UpdateValue(value, X.CellValues.Number);

    internal void UpdateValue(string value, X.CellValues type)
    {
        var stream = embeddedPackagePart.GetStream();
        var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, true);
        var xSheet = sdkSpreadsheetDocument.WorkbookPart!.Workbook.Sheets!.Elements<X.Sheet>()
            .First(xSheet => xSheet.Name == sheetName);
        var sdkWorksheetPart = (WorksheetPart)sdkSpreadsheetDocument.WorkbookPart!.GetPartById(xSheet.Id!);
        var xCells = sdkWorksheetPart.Worksheet.Descendants<X.Cell>();
        var xCell = xCells.FirstOrDefault(xCell => xCell.CellReference == address);

        if (xCell != null)
        {
            xCell.DataType = new EnumValue<X.CellValues>(type);
            xCell.CellValue = new X.CellValue(value);
        }
        else
        {
            var xWorksheet = sdkWorksheetPart.Worksheet;
            var xSheetData = xWorksheet.Elements<X.SheetData>().First();
            var rowNumberStr = Regex.Match(address, @"\d+", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value;
            var rowNumber = int.Parse(rowNumberStr, NumberStyles.Number, NumberFormatInfo.InvariantInfo);
            var xRow = xSheetData.Elements<X.Row>().First(r => r.RowIndex! == rowNumber);
            var newXCell = new X.Cell
            {
                CellReference = address,
                DataType = new EnumValue<X.CellValues>(type),
                CellValue = new X.CellValue(value)
            };

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            X.Cell? refCell = xRow.Elements<X.Cell>().FirstOrDefault(cell => string.Compare(cell.CellReference!.Value, address, true, CultureInfo.InvariantCulture) > 0);

            xRow.InsertBefore(newXCell, refCell);
        }

        sdkSpreadsheetDocument.Dispose();
        stream.Close();
    }
}