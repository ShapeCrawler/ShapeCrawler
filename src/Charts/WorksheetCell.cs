using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts;

internal sealed class WorksheetCell(EmbeddedPackagePart embeddedPackagePart, string sheetName, string address)
{
    internal void UpdateValue(string value) => this.UpdateValue(value, DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);

    internal void UpdateValue(string value, DocumentFormat.OpenXml.Spreadsheet.CellValues type)
    {
        var stream = embeddedPackagePart.GetStream();
        var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, true);
        var xSheet = sdkSpreadsheetDocument.WorkbookPart!.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .First(xSheet => xSheet.Name == sheetName);
        var sdkWorksheetPart = (WorksheetPart)sdkSpreadsheetDocument.WorkbookPart!.GetPartById(xSheet.Id!);
        var xCells = sdkWorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>();
        var xCell = xCells.FirstOrDefault(xCell => xCell.CellReference == address);

        if (xCell != null)
        {
            xCell.DataType = new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(type);
            xCell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value);
        }
        else
        {
            var xWorksheet = sdkWorksheetPart.Worksheet;
            var xSheetData = xWorksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().First();
            var rowNumberStr = Regex.Match(address, @"\d+", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value;
            var rowNumber = int.Parse(rowNumberStr, NumberStyles.Number, NumberFormatInfo.InvariantInfo);
            var xRow = xSheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().First(r => r.RowIndex! == rowNumber);
            var newXCell = new DocumentFormat.OpenXml.Spreadsheet.Cell
            {
                CellReference = address,
                DataType = new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(type),
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value)
            };

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            var refCell = xRow.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().FirstOrDefault(cell => string.Compare(cell.CellReference!.Value, address, true, CultureInfo.InvariantCulture) > 0);

            xRow.InsertBefore(newXCell, refCell);
        }

        sdkSpreadsheetDocument.Dispose();
        stream.Close();
    }
}