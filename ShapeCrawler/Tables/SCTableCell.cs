using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Placeholders;
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
            this.SdkATableCell = aTableCell;
            this.RowIndex = rowIndex;
            this.ColumnIndex = columnIndex;
            this.textBox = new ResettableLazy<SCTextBox>(this.GetTextBox);
        }

        #region Public Properties

        public ITextBox TextBox => this.textBox.Value;

        /// <inheritdoc/>
        public bool IsMergedCell => this.DefineWhetherCellIsMerged();

        public SCSlideMaster ParentSlideMaster => this.ParentTableRow.ParentTable.ParentSlideMaster;

        public IPlaceholder Placeholder => throw new System.NotImplementedException();

        #endregion Public Properties

        internal int RowIndex { get; }

        internal int ColumnIndex { get; }

        internal SCTableRow ParentTableRow { get; }

        internal A.TableCell SdkATableCell { get; init; }

        private SCTextBox GetTextBox()
        {
            A.TextBody aTextBody = this.SdkATableCell.TextBody;
            IEnumerable<A.Text> aTexts = aTextBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0)
            {
                return new SCTextBox(aTextBody, this);
            }

            return null;
        }

        private bool DefineWhetherCellIsMerged()
        {
            return this.SdkATableCell.GridSpan != null ||
                   this.SdkATableCell.RowSpan != null ||
                   this.SdkATableCell.HorizontalMerge != null ||
                   this.SdkATableCell.VerticalMerge != null;
        }

        public void ThrowIfRemoved()
        {
            if (this.isRemoved)
            {
                throw new ElementIsRemovedException("Table Cell was removed.");
            }

            this.ParentTableRow.ThrowIfRemoved();
        }


    }
}