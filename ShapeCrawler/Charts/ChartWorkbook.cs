using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts
{
    internal class ChartWorkbook //TODO: implement IDispose to correctly dispose _packagePartStream
    {
        private readonly SCChart _slideChart;
        private readonly Lazy<WorkbookPart> _workbookPart;
        private Stream _packagePartStream;
        private MemoryStream _resizableStream;

        internal ChartWorkbook(SCChart slideChart)
        {
            _slideChart = slideChart;
            _workbookPart = new Lazy<WorkbookPart>(GetWorkbookPart);
        }

        internal WorkbookPart WorkbookPart => _workbookPart.Value;

        internal void Close()
        {
            _resizableStream?.WriteTo(_packagePartStream);
            _packagePartStream?.Close();
        }

        private WorkbookPart GetWorkbookPart()
        {
            SpreadsheetDocument spreadsheetDocument;
            _packagePartStream = _slideChart.SdkChartPart.EmbeddedPackagePart.GetStream();
            if (_slideChart.ParentPresentation.Editable)
            {
                _resizableStream = new MemoryStream();
                _packagePartStream.CopyTo(_resizableStream);
                spreadsheetDocument = SpreadsheetDocument.Open(_resizableStream, true);
            }
            else
            {
                spreadsheetDocument = SpreadsheetDocument.Open(_packagePartStream, false);
            }

            _slideChart.ParentPresentation.ChartWorkbooks.Add(this);

            return spreadsheetDocument.WorkbookPart;
        }
    }
}