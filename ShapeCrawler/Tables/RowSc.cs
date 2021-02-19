using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables
{
    /// <summary>
    ///     Represents a row in a table.
    /// </summary>
    public class RowSc
    {
        private readonly Lazy<List<CellSc>> _cells;

        internal readonly A.TableRow ATableRow;
        internal readonly int Index;

        #region Constructors

        internal RowSc(TableSc table, A.TableRow aTableRow, int index)
        {
            Table = table;
            ATableRow = aTableRow;
            Index = index;

#if NETSTANDARD2_0
            _cells = new Lazy<List<CellSc>>(() => GetCells());
#else
            _cells = new Lazy<List<CellSc>>(GetCells);
#endif
        }

        #endregion Constructors

        #region Private Methods

        private List<CellSc> GetCells()
        {
            var cellList = new List<CellSc>();
            IEnumerable<A.TableCell> aTableCells = ATableRow.Elements<A.TableCell>();
            CellSc addedCell = null;

            int columnIdx = 0;
            foreach (A.TableCell aTableCell in aTableCells)
            {
                if (aTableCell.HorizontalMerge != null)
                {
                    cellList.Add(addedCell);
                }
                else if (aTableCell.VerticalMerge != null)
                {
                    int upRowIdx = Index - 1;
                    CellSc upNeighborCell = Table[upRowIdx, columnIdx];
                    cellList.Add(upNeighborCell);
                    addedCell = upNeighborCell;
                }
                else
                {
                    addedCell = new CellSc(Table, aTableCell, Index, columnIdx);
                    cellList.Add(addedCell);
                }

                columnIdx++;
            }

            return cellList;
        }

        #endregion

        #region Public Properties

        /// <summary>
        ///     Returns row's cells.
        /// </summary>
        public IReadOnlyList<CellSc> Cells => _cells.Value;

        public TableSc Table { get; }

        public long Height
        {
            get => ATableRow.Height.Value;
            set => ATableRow.Height.Value = value;
        }

        #endregion Public Properties
    }
}