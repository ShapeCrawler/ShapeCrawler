using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class ChartPoint : IChartPoint
{
    private readonly ChartPart sdkChartPart;
    private readonly C.NumericValue cNumericValue;
    private readonly string sheet;
    private readonly string address;

    internal ChartPoint(ChartPart sdkChartPart, C.NumericValue cNumericValue, string sheet, string address)
    {
        this.sdkChartPart = sdkChartPart;
        this.cNumericValue = cNumericValue;
        this.sheet = sheet;
        this.address = address;
    }

    public double Value
    {
        get => this.ParseValue();
        set => this.UpdateValue(value);
    }

    private double ParseValue()
    {
        var cachedValue = double.Parse(this.cNumericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);

        return Math.Round(cachedValue, 2);
    }

    private void UpdateValue(double value)
    {
        this.cNumericValue.Text = value.ToString(CultureInfo.InvariantCulture);

        if (this.sdkChartPart.EmbeddedPackagePart == null)
        {
            return;
        }

        var stream = this.sdkChartPart.EmbeddedPackagePart.GetStream();
        var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, true);
        
        var xSheet = sdkSpreadsheetDocument.WorkbookPart!.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == this.sheet);
        var sdkWorksheetPart = (WorksheetPart)sdkSpreadsheetDocument.WorkbookPart!.GetPartById(xSheet.Id!);
        var xCells = sdkWorksheetPart.Worksheet.Descendants<X.Cell>();

        var xCell = xCells.FirstOrDefault(xCell => xCell.CellReference == this.address);
        
        if (xCell != null)
        {
            xCell.DataType = new EnumValue<X.CellValues>(X.CellValues.Number);
            xCell.CellValue = new X.CellValue(value);
        }
        else
        {
            xSheet = sdkSpreadsheetDocument.WorkbookPart!.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == this.sheet);
            sdkWorksheetPart = (WorksheetPart)sdkSpreadsheetDocument.WorkbookPart!.GetPartById(xSheet.Id!);
            var worksheet = sdkWorksheetPart.Worksheet;
            var sheetData = worksheet.Elements<X.SheetData>().First();
            var rowNumberStr = Regex.Match(this.address, @"\d+").Value;
            var rowNumber = int.Parse(rowNumberStr, NumberStyles.Number, NumberFormatInfo.InvariantInfo);

            var row = sheetData.Elements<X.Row>().First(r => r.RowIndex! == rowNumber);
            var newXCell = new X.Cell
            {
                CellReference = this.address
            };
            newXCell.DataType = new EnumValue<X.CellValues>(X.CellValues.Number);
            newXCell.CellValue = new X.CellValue(value);

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            X.Cell? refCell = null;
            foreach (var cell in row.Elements<X.Cell>())
            {
                if (string.Compare(cell.CellReference!.Value, this.address, true, CultureInfo.InvariantCulture) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            row.InsertBefore(newXCell, refCell);
        }
        
        sdkSpreadsheetDocument.Dispose();
        stream.Close();
    }
}