using System;
using System.Linq;
using ShapeCrawler.Settings;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table row cell.
    /// </summary>
    public class Cell
    {
        #region Fields

        private TextBoxSc _textBox;

        private readonly A.TableCell _aTableCell;
        private readonly ShapeContext _spContext;

        #endregion Fields

        #region Public Properties

        /// <summary>
        /// Gets text box.
        /// </summary>
        public TextBoxSc TextBoxBox
        {
            get
            {
                if (_textBox == null)
                {
                    TryParseTxtBody();
                }

                return _textBox;
            }
        }

        public bool IsMergedCell => DefineWhetherCellIsMerged();

        #endregion Public Properties

        #region Constructors

        internal Cell(A.TableCell aTableCell)
        {
            _aTableCell = aTableCell ?? throw new ArgumentNullException(nameof(aTableCell));
        }

        #endregion Constructors

        private void TryParseTxtBody()
        {
            var aTxtBody = _aTableCell.TextBody;
            var aTexts = aTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                _textBox = new TextBoxSc(_spContext, aTxtBody);
            }
        }

        private bool DefineWhetherCellIsMerged()
        {
            return _aTableCell.GridSpan != null ||
                   _aTableCell.RowSpan != null ||
                   _aTableCell.HorizontalMerge != null ||
                   _aTableCell.VerticalMerge != null;
        }
    }
}
