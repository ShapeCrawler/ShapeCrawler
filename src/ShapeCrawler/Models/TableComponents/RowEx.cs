using System;
using System.Collections.Generic;
using System.Linq;
using SlideDotNet.Models.Settings;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.TableComponents
{
    /// <summary>
    /// Represents a table's row.
    /// </summary>
    public class RowEx
    {
        #region Fields

        private List<CellEx> _cells;
        private readonly A.TableRow _sdkTblRow;
        private readonly IShapeContext _spContext;

        #endregion

        #region Properties

        /// <summary>
        /// Returns row's cells.
        /// </summary>
        /// TODO: use custom collection
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

        public RowEx(A.TableRow xmlRow, IShapeContext spContext)
        {
            _sdkTblRow = xmlRow ?? throw new ArgumentNullException(nameof(xmlRow));
            _spContext = spContext ?? throw new ArgumentNullException(nameof(spContext));
        }

        #endregion

        #region Private Methods

        private void ParseCells()
        {
            var xmlCells = _sdkTblRow.Elements<A.TableCell>();
            _cells = new List<CellEx>(xmlCells.Count());
            foreach (var c in xmlCells)
            {
                _cells.Add(new CellEx(c, _spContext));
            }
        }

        #endregion
    }
}
