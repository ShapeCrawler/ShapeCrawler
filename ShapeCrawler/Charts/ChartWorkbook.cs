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
        private SpreadsheetDocument spreadsheetDocument;
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


            this.spreadsheetDocument.Close();

            this.closed = true;
        }

        private WorkbookPart GetWorkbookPart()
        {
            this.packagePartStream = this.chart.SdkChartPart.EmbeddedPackagePart.GetStream();
            this.spreadsheetDocument = SpreadsheetDocument.Open(packagePartStream, true);
            this.chart.ParentPresentation.ChartWorkbooks.Add(this);

            return this.spreadsheetDocument.WorkbookPart;
        }
    }
}