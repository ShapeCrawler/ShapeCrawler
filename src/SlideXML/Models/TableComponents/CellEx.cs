using System.Linq;
using LogicNull.Utilities;
using SlideXML.Models.Settings;
using SlideXML.Models.TextBody;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Models.TableComponents
{
    /// <summary>
    /// Represents a row's cell.
    /// </summary>
    public class CellEx
    {
        #region Fields

        private TextBodyEx _textBody;

        private readonly A.TableCell _xmlCell;
        private readonly ElementSettings _elSettings;

        #endregion

        #region Properties

        /// <summary>
        /// Returns <see cref="TextBodyEx"/> instance or null if the cell does not contain a text.
        /// </summary>
        public TextBodyEx TextBody
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

        public CellEx(A.TableCell xmlCell, ElementSettings elSettings)
        {
            Check.NotNull(xmlCell, nameof(xmlCell));
            Check.NotNull(elSettings, nameof(elSettings));
            _xmlCell = xmlCell;
            _elSettings = elSettings;
        }

        #endregion

        private void TryParseTxtBody()
        {
            var aTxtBody = _xmlCell.TextBody;
            var aTexts = aTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                _textBody = new TextBodyEx(_elSettings, aTxtBody);
            }
        }
    }
}
