using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table's column.
    /// </summary>
    public class Column
    {
        private readonly GridColumn _aGridColumn;

        public Column(GridColumn aGridColumn)
        {
            _aGridColumn = aGridColumn;
        }

        public long Width => _aGridColumn.Width.Value;
    }
}
