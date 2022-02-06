using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts
{
    internal class ChartWorkbook // TODO: implement IDispose to correctly dispose _packagePartStream
    {
        private readonly SCChart chart;
        private readonly Lazy<SpreadsheetDocument> spreadsheetDocument;
        private Stream embeddedPackagePartStream;

        private EmbeddedPackagePart embeddedPackagePart;

        private bool closed;

        internal ChartWorkbook(SCChart chart, EmbeddedPackagePart embeddedPackagePart)
        {
            this.chart = chart;
            this.embeddedPackagePart = embeddedPackagePart;
            this.spreadsheetDocument = new Lazy<SpreadsheetDocument>(this.GetSpreadsheetDocument);
        }

        internal WorkbookPart WorkbookPart => this.spreadsheetDocument.Value.WorkbookPart;

        internal byte[] ByteArray => this.GetByteArray();

        internal void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.spreadsheetDocument.Value?.Close();
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
        
        internal X.Cell? GetXCellOrDefault(string sheetName, string cellAddress)
        {
            var chartSheet = this.WorkbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == sheetName);
            var worksheetPart = (WorksheetPart)this.WorkbookPart.GetPartById(chartSheet.Id!);
            var sheetXCells = worksheetPart.Worksheet.Descendants<X.Cell>();

            return sheetXCells.FirstOrDefault(xCell => xCell.CellReference == cellAddress);
        }

        private SpreadsheetDocument GetSpreadsheetDocument()
        {
            this.embeddedPackagePartStream = this.embeddedPackagePart.GetStream();
            var spreadsheetDocument = SpreadsheetDocument.Open(this.embeddedPackagePartStream, this.chart.PresentationInternal.Editable);
            this.chart.PresentationInternal.ChartWorkbooks.Add(this);

            return spreadsheetDocument;
        }

        private byte[] GetByteArray()
        {
            var mStream = new MemoryStream();
            this.spreadsheetDocument.Value.Clone(mStream);

            return mStream.ToArray();
        }
    }
}