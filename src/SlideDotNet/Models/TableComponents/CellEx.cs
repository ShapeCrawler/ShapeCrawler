using System;
using System.Linq;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.TextBody;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.TableComponents
{
    /// <summary>
    /// Represents a row's cell.
    /// </summary>
    public class CellEx
    {
        #region Fields

        private TextFrame _textBody;

        private readonly A.TableCell _xmlCell;
        private readonly IShapeContext _spContext;

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

        public CellEx(A.TableCell xmlCell, IShapeContext spContext)
        {
            _xmlCell = xmlCell ?? throw new ArgumentNullException(nameof(xmlCell));
            _spContext = spContext ?? throw new ArgumentNullException(nameof(spContext));
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
