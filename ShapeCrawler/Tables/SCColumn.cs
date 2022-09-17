using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a table's column.
    /// </summary>
    public class SCColumn
    {
        internal SCColumn(A.GridColumn aGridColumn)
        {
            this.AGridColumn = aGridColumn;
        }
        
        /// <summary>
        ///     Gets or sets cell width.
        /// </summary>
        public long Width
        {
            get => this.AGridColumn.Width.Value;
            set => this.AGridColumn.Width.Value = value;
        }

        internal A.GridColumn AGridColumn { get; init; }
    }
}