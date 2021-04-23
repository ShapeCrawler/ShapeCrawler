using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a table's column.
    /// </summary>
    public class Column
    {
        internal Column(A.GridColumn aGridColumn)
        {
            this.AGridColumn = aGridColumn;
        }

        internal A.GridColumn AGridColumn { get; init; }

        public long Width
        {
            get => this.AGridColumn.Width.Value;
            set => this.AGridColumn.Width.Value = value;
        }
    }
}