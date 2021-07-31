using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts
{
    internal class ChartWorkbook // TODO: implement IDispose to correctly dispose _packagePartStream
    {
        private readonly SCChart chart;
        private readonly Lazy<WorkbookPart> sdkWorkbookPart;
        private Stream packagePartStream;
        private MemoryStream resizableStream;
        private bool closed;

        internal ChartWorkbook(SCChart chart)
        {
            this.chart = chart;
            this.sdkWorkbookPart = new Lazy<WorkbookPart>(this.GetWorkbookPart);
        }

        internal WorkbookPart WorkbookPart => this.sdkWorkbookPart.Value;

        internal void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.resizableStream?.WriteTo(this.packagePartStream);
            this.packagePartStream?.Close();

            this.closed = true;
        }

        private WorkbookPart GetWorkbookPart()
        {
            SpreadsheetDocument spreadsheetDocument;
            this.packagePartStream = this.chart.SdkChartPart.EmbeddedPackagePart.GetStream();
            if (this.chart.ParentPresentation.Editable)
            {
                this.resizableStream = new MemoryStream();
                this.packagePartStream.CopyTo(this.resizableStream);
                spreadsheetDocument = SpreadsheetDocument.Open(this.resizableStream, true);
            }
            else
            {
                spreadsheetDocument = SpreadsheetDocument.Open(this.packagePartStream, false);
            }

            this.chart.ParentPresentation.ChartWorkbooks.Add(this);

            return spreadsheetDocument.WorkbookPart;
        }
    }
}