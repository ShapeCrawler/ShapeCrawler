using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Charts
{
    internal class ChartWorkbook
    {
        private readonly SlideChart _slideChart;
        private readonly Lazy<WorkbookPart> _workbookPart;
        private Stream _packagePartStream;
        private MemoryStream _resizableStream;
        internal WorkbookPart WorkbookPart => _workbookPart.Value;

        internal ChartWorkbook(SlideChart slideChart)
        {
            _slideChart = slideChart;
            _workbookPart = new Lazy<WorkbookPart>(() => GetWorkbookPart());
        }

        internal void Close()
        {
            _resizableStream?.WriteTo(_packagePartStream);
            _packagePartStream?.Close();
        }

        private WorkbookPart GetWorkbookPart() // TODO: set using statements
        {
            SpreadsheetDocument spreadsheetDocument;
            _packagePartStream = _slideChart.ChartPart.EmbeddedPackagePart.GetStream();
            if (_slideChart.Presentation.Editable)
            {
                _resizableStream = new MemoryStream();
                _packagePartStream.CopyTo(_resizableStream);
                spreadsheetDocument = SpreadsheetDocument.Open(_resizableStream, true);
            }
            else
            {
                spreadsheetDocument = SpreadsheetDocument.Open(_packagePartStream, false);
            }

            _slideChart.Presentation.ChartWorkbooks.Add(this);

            return spreadsheetDocument.WorkbookPart;
        }
    }
}