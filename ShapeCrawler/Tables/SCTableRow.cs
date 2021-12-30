using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a row in a table.
    /// </summary>
    public class SCTableRow // TODO: extract interface
    {
        private readonly Lazy<List<SCTableCell>> cells;
        private readonly int index;
        private readonly bool isRemoved;

        internal readonly A.TableRow ATableRow;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCTableRow"/> class.
        /// </summary>
        internal SCTableRow(SlideTable table, A.TableRow aTableRow, int index)
        {
            this.ParentTable = table;
            this.ATableRow = aTableRow;
            this.index = index;

#if NETSTANDARD2_0
            cells = new Lazy<List<SCTableCell>>(() => GetCells());
#else
            this.cells = new Lazy<List<SCTableCell>>(this.GetCells);
#endif
        }

        /// <summary>
        ///     Gets row's cells.
        /// </summary>
        public IReadOnlyList<ITableCell> Cells => this.cells.Value;

        public long Height
        {
            get => this.ATableRow.Height.Value;
            set => this.ATableRow.Height.Value = value;
        }

        internal SlideTable ParentTable { get; }

        #region Private Methods

        private List<SCTableCell> GetCells()
        {
            var cellList = new List<SCTableCell>();
            IEnumerable<A.TableCell> aTableCells = this.ATableRow.Elements<A.TableCell>();
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
                    int upRowIdx = index - 1;
                    SCTableCell upNeighborScCell = (SCTableCell) ParentTable[upRowIdx, columnIdx];
                    cellList.Add(upNeighborScCell);
                    addedScCell = upNeighborScCell;
                }
                else
                {
                    addedScCell = new SCTableCell(this, aTableCell, index, columnIdx);
                    cellList.Add(addedScCell);
                }

                columnIdx++;
            }

            return cellList;
        }

        internal void ThrowIfRemoved()
        {
            if (this.isRemoved)
            {
                throw new ElementIsRemovedException("Table Row was removed.");
            }

            this.ParentTable.ThrowIfRemoved();
        }

        #endregion
    }
}