using System.Collections.Generic;
using System.Linq;
using SlideXML.Models.Settings;
using SlideXML.Validation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Models.TableComponents
{
    /// <summary>
    /// Represents a table's row.
    /// </summary>
    public class RowEx
    {
        #region Fields

        private List<CellEx> _cells;
        private readonly A.TableRow _xmlRow;
        private readonly ElementSettings _elSettings;

        #endregion

        #region Properties

        /// <summary>
        /// Returns row's cells.
        /// </summary>
        public IList<CellEx> Cells {
            get
            {
                if (_cells == null)
                {
                    ParseCells();
                }

                return _cells;
            }
        }

        #endregion

        #region Constructors

        public RowEx(A.TableRow xmlRow, ElementSettings elSettings)
        {
            Check.NotNull(xmlRow, nameof(xmlRow));
            Check.NotNull(elSettings, nameof(elSettings));
            _xmlRow = xmlRow;
            _elSettings = elSettings;
        }

        #endregion

        #region Private Methods

        private void ParseCells()
        {
            var xmlCells = _xmlRow.Elements<A.TableCell>();
            _cells = new List<CellEx>(xmlCells.Count());
            foreach (var c in xmlCells)
            {
                _cells.Add(new CellEx(c, _elSettings));
            }
        }

        #endregion
    }
}
