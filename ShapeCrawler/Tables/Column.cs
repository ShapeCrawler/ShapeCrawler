using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    ///     Represents a table's column.
    /// </summary>
    public class Column
    {
        internal Column(A.GridColumn aGridColumn)
        {
            AGridColumn = aGridColumn;
        }

        internal A.GridColumn AGridColumn { get; init; }

        public long Width
        {
            get => AGridColumn.Width.Value;
            set => AGridColumn.Width.Value = value;
        }
    }
}