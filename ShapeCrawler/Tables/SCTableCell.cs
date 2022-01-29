using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    internal class SCTableCell : ITableCell, ITextBoxContainer
    {
        private readonly ResettableLazy<SCTextBox> textBox;
        private readonly bool isRemoved;

        internal SCTableCell(SCTableRow parentTableRow, A.TableCell aTableCell, int rowIndex, int columnIndex)
        {
            this.ParentTableRow = parentTableRow;
            this.ATableCell = aTableCell;
            this.RowIndex = rowIndex;
            this.ColumnIndex = columnIndex;
            this.textBox = new ResettableLazy<SCTextBox>(this.GetTextBox);
        }

        public bool IsMergedCell => this.DefineWhetherCellIsMerged();

        public SCSlideMaster SlideMasterInternal => this.ParentTableRow.ParentTable.SlideMasterInternal;

        public IPlaceholder Placeholder => throw new System.NotImplementedException();

        public IShape Shape => this.ParentTableRow.ParentTable;

        public ITextBox TextBox => this.textBox.Value;

        internal A.TableCell ATableCell { get; init; }

        internal int RowIndex { get; }

        internal int ColumnIndex { get; }

        private SCTableRow ParentTableRow { get; }

        public void ThrowIfRemoved()
        {
            if (this.isRemoved)
            {
                throw new ElementIsRemovedException("Table Cell was removed.");
            }

            this.ParentTableRow.ThrowIfRemoved();
        }

        private SCTextBox GetTextBox()
        {
            A.TextBody aTextBody = this.ATableCell.TextBody;
            IEnumerable<A.Text> aTexts = aTextBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0)
            {
                return new SCTextBox(aTextBody, this);
            }

            return null;
        }

        private bool DefineWhetherCellIsMerged()
        {
            return this.ATableCell.GridSpan != null ||
                   this.ATableCell.RowSpan != null ||
                   this.ATableCell.HorizontalMerge != null ||
                   this.ATableCell.VerticalMerge != null;
        }

    }
}