using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts
{
    internal class ChartWorkbook // TODO: implement IDispose to correctly dispose _packagePartStream
    {
        private readonly SCChart chart;
        private readonly Lazy<WorkbookPart> workbookPart;
        private Stream embeddedPackagePartStream;
        private SpreadsheetDocument spreadsheetDocument;
        private bool closed;

        internal ChartWorkbook(SCChart chart)
        {
            this.chart = chart;
            this.workbookPart = new Lazy<WorkbookPart>(this.GetWorkbookPart);
        }

        internal WorkbookPart WorkbookPart => this.workbookPart.Value;

        internal byte[] ByteArray => this.GetByteArray();

        internal void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.spreadsheetDocument?.Close();
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

        private WorkbookPart GetWorkbookPart()
        {
            this.embeddedPackagePartStream = this.chart.SdkChartPart.EmbeddedPackagePart.GetStream();
            this.spreadsheetDocument = SpreadsheetDocument.Open(this.embeddedPackagePartStream, this.chart.ParentPresentationInternal.Editable);
            this.chart.ParentPresentationInternal.ChartWorkbooks.Add(this);

            return this.spreadsheetDocument.WorkbookPart;
        }

        private byte[] GetByteArray()
        {
            var mStream = new MemoryStream();
            this.spreadsheetDocument.Clone(mStream);

            return mStream.ToArray();
        }
    }
}