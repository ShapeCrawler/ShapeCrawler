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
        private readonly Lazy<List<SCTableCell>> _cells;
        internal readonly A.TableRow ATableRow;
        internal readonly int Index;

        #region Constructors

        internal RowSc(SlideTable table, A.TableRow aTableRow, int index)
        {
            Table = table;
            ATableRow = aTableRow;
            Index = index;

#if NETSTANDARD2_0
            _cells = new Lazy<List<CellSc>>(() => GetCells());
#else
            _cells = new Lazy<List<SCTableCell>>(GetCells);
#endif
        }

        #endregion Constructors

        internal SlideTable Table { get; }

        #region Private Methods

        private List<SCTableCell> GetCells()
        {
            var cellList = new List<SCTableCell>();
            IEnumerable<A.TableCell> aTableCells = ATableRow.Elements<A.TableCell>();
            SCTableCell addedScCell = null;

            int columnIdx = 0;
            foreach (A.TableCell aTableCell in aTableCells)
            {
                if (aTableCell.HorizontalMerge != null)
                {
                    cellList.Add(addedScCell);
                }
                else if (aTableCell.VerticalMerge != null)
                {
                    int upRowIdx = Index - 1;
                    SCTableCell upNeighborScCell = (SCTableCell)Table[upRowIdx, columnIdx];
                    cellList.Add(upNeighborScCell);
                    addedScCell = upNeighborScCell;
                }
                else
                {
                    addedScCell = new SCTableCell(Table, aTableCell, Index, columnIdx);
                    cellList.Add(addedScCell);
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
        public IReadOnlyList<ITableCell> Cells => _cells.Value;

        public long Height
        {
            get => ATableRow.Height.Value;
            set => ATableRow.Height.Value = value;
        }

        #endregion Public Properties
    }
}