using System;
using System.Linq;
using ShapeCrawler.Settings;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table row cell.
    /// </summary>
    public class CellSc
    {
        #region Fields

        private TextSc _text;

        private readonly A.TableCell _aTableCell;
        private readonly ShapeContext _spContext;

        #endregion

        #region Properties

        /// <summary>
        /// Gets text frame of a cell.
        /// </summary>
        public TextSc Text
        {
            get
            {
                if (_text == null)
                {
                    TryParseTxtBody();
                }

                return _text;
            }
        }

        public bool IsMergedCell => DefineWhetherCellIsMerged();

        private bool DefineWhetherCellIsMerged()
        {
            return _aTableCell.GridSpan != null ||
                   _aTableCell.RowSpan != null ||
                   _aTableCell.HorizontalMerge != null ||
                   _aTableCell.VerticalMerge != null;
        }

        #endregion

        #region Constructors

        public CellSc(A.TableCell xmlCell)
        {
            _aTableCell = xmlCell ?? throw new ArgumentNullException(nameof(xmlCell));
        }

        #endregion

        private void TryParseTxtBody()
        {
            var aTxtBody = _aTableCell.TextBody;
            var aTexts = aTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                _text = new TextSc(_spContext, aTxtBody);
            }
        }
    }
}
