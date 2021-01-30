using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a row in a table.
    /// </summary>
    public class RowSc
    {
        #region Fields

        private List<CellSc> _cells;
        private readonly A.TableRow _aTableRow;

        #endregion

        #region Properties

        /// <summary>
        /// Returns row's cells.
        /// </summary>
        /// TODO: use custom collection
        public IReadOnlyList<CellSc> Cells {
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

        public RowSc(A.TableRow xmlRow)
        {
            _aTableRow = xmlRow ?? throw new ArgumentNullException(nameof(xmlRow));
        }

        #endregion

        #region Private Methods

        private void ParseCells()
        {
            var xmlCells = _aTableRow.Elements<A.TableCell>();
            _cells = new List<CellSc>(xmlCells.Count());
            foreach (var c in xmlCells)
            {
                _cells.Add(new CellSc(c));
            }
        }

        #endregion
    }
}
