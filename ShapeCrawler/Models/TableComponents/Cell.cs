using System;
using System.Linq;
using ShapeCrawler.Models.TextShape;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Models.TableComponents
{
    /// <summary>
    /// Represents cell of table row.
    /// </summary>
    public class Cell
    {
        #region Fields

        private TextFrame _textBody;

        private readonly A.TableCell _xmlCell;
        private readonly ShapeContext _spContext;

        #endregion

        #region Properties

        /// <summary>
        /// Returns <see cref="TextFrame"/> instance or null if the cell does not contain a text.
        /// </summary>
        public TextFrame TextBody
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

        public Cell(A.TableCell xmlCell)
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
                _textBody = new TextFrame(_spContext, aTxtBody);
            }
        }
    }
}
