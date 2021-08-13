using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts
{
    internal class ChartWorkbook // TODO: implement IDispose to correctly dispose _packagePartStream
    {
        private readonly SCChart chart;
        private readonly Lazy<WorkbookPart> sdkWorkbookPart;
        private SpreadsheetDocument spreadsheetDocument;
        private bool closed;

        public ChartWorkbook(SCChart chart)
        {
            this.chart = chart;
            this.sdkWorkbookPart = new Lazy<WorkbookPart>(this.GetWorkbookPart);
        }

        public WorkbookPart WorkbookPart => this.sdkWorkbookPart.Value;

        public byte[] ByteArray => this.GetByteArray();

        public void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.spreadsheetDocument?.Close();
            this.closed = true;
        }

        private WorkbookPart GetWorkbookPart()
        {
            this.spreadsheetDocument = SpreadsheetDocument.Open(this.chart.SdkChartPart.EmbeddedPackagePart.GetStream(), this.chart.ParentPresentation.Editable);
            this.chart.ParentPresentation.ChartWorkbooks.Add(this);

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