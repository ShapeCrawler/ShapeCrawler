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
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrwaler")]
    public class SCTableRow // TODO: extract interface
    {
        private readonly Lazy<List<SCTableCell>> cells;
        internal readonly A.TableRow SdkATableRow;
        internal readonly int Index;
        private readonly bool isRemoved;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCTableRow"/> class.
        /// </summary>
        internal SCTableRow(SlideTable table, A.TableRow aTableRow, int index)
        {
            this.ParentTable = table;
            this.SdkATableRow = aTableRow;
            this.Index = index;

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
            get => this.SdkATableRow.Height.Value;
            set => this.SdkATableRow.Height.Value = value;
        }

        internal SlideTable ParentTable { get; }

        #region Private Methods

        private List<SCTableCell> GetCells()
        {
            var cellList = new List<SCTableCell>();
            IEnumerable<A.TableCell> aTableCells = this.SdkATableRow.Elements<A.TableCell>();
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
                    SCTableCell upNeighborScCell = (SCTableCell) ParentTable[upRowIdx, columnIdx];
                    cellList.Add(upNeighborScCell);
                    addedScCell = upNeighborScCell;
                }
                else
                {
                    addedScCell = new SCTableCell(this, aTableCell, Index, columnIdx);
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

        #region Public Properties



        #endregion Public Properties
    }
}