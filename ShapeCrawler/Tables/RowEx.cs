using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table's row.
    /// </summary>
    public class RowEx
    {
        #region Fields

        private List<CellSc> _cells;
        private readonly A.TableRow _sdkTblRow;

        #endregion

        #region Properties

        /// <summary>
        /// Returns row's cells.
        /// </summary>
        /// TODO: use custom collection
        public IList<CellSc> Cells {
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

        public RowEx(A.TableRow xmlRow)
        {
            _sdkTblRow = xmlRow ?? throw new ArgumentNullException(nameof(xmlRow));
        }

        #endregion

        #region Private Methods

        private void ParseCells()
        {
            var xmlCells = _sdkTblRow.Elements<A.TableCell>();
            _cells = new List<CellSc>(xmlCells.Count());
            foreach (var c in xmlCells)
            {
                _cells.Add(new CellSc(c));
            }
        }

        #endregion
    }
}
