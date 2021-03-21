using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <summary>
    ///     Represents a cell in a table.
    /// </summary>
    public class CellSc //TODO: make it internal
    {
        #region Constructors

        internal CellSc(SlideTable table, A.TableCell aTableCell, int rowIdx, int columnIdx)
        {
            Table = table;
            ATableCell = aTableCell;
            RowIndex = rowIdx;
            ColumnIndex = columnIdx;
            _textBox = new ResettableLazy<SCTextBox>(() => GetTextBox());
        }

        #endregion Constructors

        #region Fields

        private readonly ResettableLazy<SCTextBox> _textBox;

        internal int RowIndex { get; }
        internal int ColumnIndex { get; }
        internal SlideTable Table { get; }
        internal A.TableCell ATableCell { get; init; }

        #endregion Fields

        #region Public Properties

        /// <summary>
        ///     Gets text box.
        /// </summary>
        public ITextBox TextBox => _textBox.Value;

        public bool IsMergedCell => DefineWhetherCellIsMerged();

        #endregion Public Properties

        #region Private Methods

        private SCTextBox GetTextBox()
        {
            var aTxtBody = ATableCell.TextBody;
            var aTexts = aTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0
            ) // at least one of <a:t> element contain text
            {
                return new SCTextBox(aTxtBody);
            }

            return null;
        }

        private bool DefineWhetherCellIsMerged()
        {
            return ATableCell.GridSpan != null ||
                   ATableCell.RowSpan != null ||
                   ATableCell.HorizontalMerge != null ||
                   ATableCell.VerticalMerge != null;
        }

        #endregion Private Methods
    }
}