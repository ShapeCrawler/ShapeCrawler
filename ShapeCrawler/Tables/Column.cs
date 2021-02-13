using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table's column.
    /// </summary>
    public class Column
    {
        internal GridColumn AGridColumn { get; init; }

        internal Column(A.GridColumn aGridColumn)
        {
            AGridColumn = aGridColumn;
        }

        public long Width
        {
            get => AGridColumn.Width.Value;
            set => AGridColumn.Width.Value = value;
        }
    }
}
