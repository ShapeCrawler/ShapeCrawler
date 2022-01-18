using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

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

        public void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.spreadsheetDocument?.Close();
            this.embeddedPackagePartStream?.Close();
            this.closed = true;
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