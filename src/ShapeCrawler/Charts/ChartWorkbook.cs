using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class ChartWorkbook // TODO: implement IDispose to correctly dispose _packagePartStream
{
    private readonly SCChart chart;
    private readonly EmbeddedPackagePart embeddedPackagePart;
    private Stream? embeddedPackagePartStream;
    private bool closed;

    internal ChartWorkbook(SCChart chart, EmbeddedPackagePart embeddedPackagePart)
    {
        this.chart = chart;
        this.embeddedPackagePart = embeddedPackagePart;
        this.SpreadsheetDocument = new Lazy<SpreadsheetDocument>(this.GetSpreadsheetDocument);
    }

    internal WorkbookPart WorkbookPart => this.SpreadsheetDocument.Value.WorkbookPart!;

    internal byte[] BinaryData => this.GetByteArray();

    internal Lazy<SpreadsheetDocument> SpreadsheetDocument { get; }

    internal void Save()
    {
        this.SpreadsheetDocument.Value.Save();
    }

    internal void Close()
    {
        if (this.closed)
        {
            return;
        }

        this.SpreadsheetDocument.Value?.Close();
        this.embeddedPackagePartStream?.Close();
        this.closed = true;
    }

    internal X.Cell GetXCell(string sheetName, string cellAddress)
    {
        var chartSheet = this.WorkbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == sheetName);
        var worksheetPart = (WorksheetPart)this.WorkbookPart.GetPartById(chartSheet.Id!);
        var sheetXCells = worksheetPart.Worksheet.Descendants<X.Cell>();

        return sheetXCells.First(xCell => xCell.CellReference == cellAddress);
    }

    internal void UpdateCell(string sheetName, string cellReference, double value)
    {
        var xCell = this.GetXCellOrDefault(sheetName, cellReference);
        if (xCell != null)
        {
            xCell.DataType = new EnumValue<X.CellValues>(X.CellValues.Number);
            xCell.CellValue = new X.CellValue(value);
        }
        else
        {
            var chartSheet = this.WorkbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == sheetName);
            var worksheetPart = (WorksheetPart)this.WorkbookPart.GetPartById(chartSheet.Id!);
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.Elements<X.SheetData>().First();
            var rowNumberStr = Regex.Match(cellReference, @"\d+").Value;
            var rowNumber = int.Parse(rowNumberStr, NumberStyles.Number, NumberFormatInfo.InvariantInfo);

            var row = sheetData.Elements<X.Row>().First(r => r.RowIndex! == rowNumber);
            var newXCell = new X.Cell
            {
                CellReference = cellReference
            };
            newXCell.DataType = new EnumValue<X.CellValues>(X.CellValues.Number);
            newXCell.CellValue = new X.CellValue(value);

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            X.Cell? refCell = null;
            foreach (var cell in row.Elements<X.Cell>())
            {
                if (string.Compare(cell.CellReference!.Value, cellReference, true, CultureInfo.InvariantCulture) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            row.InsertBefore(newXCell, refCell);
        }
    }

    private X.Cell? GetXCellOrDefault(string sheetName, string cellAddress)
    {
        var chartSheet = this.WorkbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == sheetName);
        var worksheetPart = (WorksheetPart)this.WorkbookPart.GetPartById(chartSheet.Id!);
        var sheetXCells = worksheetPart.Worksheet.Descendants<X.Cell>();

        return sheetXCells.FirstOrDefault(xCell => xCell.CellReference == cellAddress);
    }

    private SpreadsheetDocument GetSpreadsheetDocument()
    {
        this.embeddedPackagePartStream = this.embeddedPackagePart.GetStream();
        var spreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(this.embeddedPackagePartStream, true);
        this.chart.SlideBase.PresentationInternal.ChartWorkbooks.Add(this);

        return spreadsheetDocument;
    }

    private byte[] GetByteArray()
    {
        var mStream = new MemoryStream();
        this.SpreadsheetDocument.Value.Clone(mStream);

        return mStream.ToArray();
    }
}