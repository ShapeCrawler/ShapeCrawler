using System;
using System.Linq;
using ShapeCrawler.Models.TextShape;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table row cell.
    /// </summary>
    public class CellSc
    {
        #region Fields

        private TextSc _textBody;

        private readonly A.TableCell _xmlCell;
        private readonly ShapeContext _spContext;

        #endregion

        #region Properties

        /// <summary>
        /// Gets text frame of a cell.
        /// </summary>
        public TextSc TextFrame
        {
            get
            {
                if (_textBody == null)
                {
                    TryParseTxtBody();
                }

                return _textBody;
            }
        }

        #endregion

        #region Constructors

        public CellSc(A.TableCell xmlCell)
        {
            _xmlCell = xmlCell ?? throw new ArgumentNullException(nameof(xmlCell));
        }

        #endregion

        private void TryParseTxtBody()
        {
            var aTxtBody = _xmlCell.TextBody;
            var aTexts = aTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                _textBody = new TextSc(_spContext, aTxtBody);
            }
        }
    }
}
