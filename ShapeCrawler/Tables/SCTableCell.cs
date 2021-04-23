using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <inheritdoc cref="ITable"/>
    internal class SCTableCell : ITableCell, ITextBoxContainer
    {
        private readonly ResettableLazy<SCTextBox> textBox;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCTableCell"/> class.
        /// </summary>
        internal SCTableCell(SCTableRow parentTableRow, A.TableCell aTableCell, int rowIndex, int columnIndex)
        {
            this.ParentTableRow = parentTableRow;
            this.ATableCell = aTableCell;
            this.RowIndex = rowIndex;
            this.ColumnIndex = columnIndex;
            this.textBox = new ResettableLazy<SCTextBox>(this.GetTextBox);
        }

        #region Public Properties

        /// <inheritdoc/>
        public ITextBox TextBox => this.textBox.Value;

        /// <inheritdoc/>
        public bool IsMergedCell => this.DefineWhetherCellIsMerged();

        #endregion Public Properties

        internal int RowIndex { get; }

        internal int ColumnIndex { get; }

        internal SCTableRow ParentTableRow { get; }

        internal A.TableCell ATableCell { get; init; }

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

        public void ThrowIfRemoved()
        {
            throw new System.NotImplementedException();
        }
    }
}