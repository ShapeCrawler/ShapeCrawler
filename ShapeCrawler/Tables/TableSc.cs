using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable SuggestVarOrType_BuiltInTypes

namespace ShapeCrawler.Tables
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableSc
    {
        #region Fields

        private readonly P.GraphicFrame _pGraphicFrame;
        private readonly ResettableLazy<RowCollection> _rowCollection;

        #endregion Fields
        
        internal ShapeSc Shape { get; set; }
        internal A.Table ATable => _pGraphicFrame.GetATable();

        #region Public Properties

        public IReadOnlyList<Column> Columns => GetColumnList(); //TODO: make lazy
        public RowCollection Rows => _rowCollection.Value;
        public CellSc this[int rowIndex, int columnIndex] => Rows[rowIndex].Cells[columnIndex];

        #endregion Public Properties

        #region Constructors

        internal TableSc(P.GraphicFrame pGraphicFrame)
        {
            _pGraphicFrame = pGraphicFrame;
            _rowCollection = new ResettableLazy<RowCollection>(() => RowCollection.Create(this, _pGraphicFrame));
        }

        #endregion Constructors

        #region Private Methods

        private IReadOnlyList<Column> GetColumnList()
        {
            IEnumerable<A.GridColumn> aGridColumns = ATable.TableGrid.Elements<A.GridColumn>();
            var columnList = new List<Column>(aGridColumns.Count());
            columnList.AddRange(aGridColumns.Select(aGridColumn => new Column(aGridColumn)));
            
            return columnList;
        }

        #endregion Private Methods

#if DEBUG
        public void MergeCells(CellSc cell1, CellSc cell2) // TODO: Optimize method
        {
            if (CannotBeMerged(cell1, cell2))
            {
                return;
            }

            int minRowIndex = cell1.RowIndex < cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
            int maxRowIndex = cell1.RowIndex > cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
            int minColIndex = cell1.ColumnIndex < cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;
            int maxColIndex = cell1.ColumnIndex > cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;

            // Horizontal merging
            List<A.TableRow> aTableRowList = cell1.Table.ATable.Elements<A.TableRow>().ToList();
            if (minColIndex != maxColIndex)
            {
                int horizontalMergingCount = maxColIndex - minColIndex + 1;
                for (int rowIdx = minRowIndex; rowIdx <= maxRowIndex; rowIdx++)
                {
                    A.TableCell[] rowATblCells = aTableRowList[rowIdx].Elements<A.TableCell>().ToArray();
                    A.TableCell firstMergingCell = rowATblCells[minColIndex];
                    firstMergingCell.GridSpan = new Int32Value(horizontalMergingCount);
                    Span<A.TableCell> nextMergingCells = new Span<A.TableCell>(rowATblCells, minColIndex + 1, horizontalMergingCount - 1);
                    foreach (A.TableCell aTblCell in nextMergingCells)
                    {
                        aTblCell.HorizontalMerge = new BooleanValue(true);
                    }
                }
            }

            // Vertical merging
            if (minRowIndex != maxRowIndex)
            {
                int verticalMergingCount = maxRowIndex - minRowIndex + 1;
                foreach (A.TableCell aTblCell in aTableRowList[minRowIndex].Elements<A.TableCell>().Skip(minColIndex).Take(maxColIndex + 1))
                {
                    aTblCell.RowSpan = new Int32Value(verticalMergingCount);
                }
                foreach (A.TableRow aTblRow in aTableRowList.Skip(minRowIndex + 1).Take(maxRowIndex))
                {
                    foreach (A.TableCell aTblCell in aTblRow.Elements<A.TableCell>().Take(maxColIndex + 1))
                    {
                        aTblCell.VerticalMerge = new BooleanValue(true);
                    }
                }
            }

            // Delete columns
            for (int colIdx = 0; colIdx < Columns.Count; )
            {
                int? gridSpan = Rows[0].Cells[colIdx].ATableCell.GridSpan?.Value;
                if (gridSpan > 1 && Rows.All(r => r.Cells[colIdx].ATableCell.GridSpan?.Value == gridSpan))
                {
                    int deleteColumnCount = gridSpan.Value - 1;
                    
                    // Delete a:gridCol elements
                    foreach (Column column in Columns.Skip(colIdx).Take(deleteColumnCount))
                    {
                        column.AGridColumn.Remove();
                        Columns[colIdx].Width += column.AGridColumn.Width; // append width of deleting column to merged column
                    }

                    // Delete a:tc elements
                    foreach (A.TableRow aTblRow in aTableRowList)
                    {
                        foreach (A.TableCell aTblCell in aTblRow.Elements<A.TableCell>().Skip(colIdx).Take(deleteColumnCount))
                        {
                            aTblCell.Remove();
                        }
                    }

                    colIdx += gridSpan.Value;
                    continue;
                }

                colIdx++;
            }

            _rowCollection.Reset();
        }

        private static bool CannotBeMerged(CellSc cell1, CellSc cell2)
        {
            if (cell1 == cell2)
            {
                // The cells are already merged
                return true;
            }

            if (cell1.Table != cell2.Table)
            {
                throw new ShapeCrawlerException("Specified cells are from different tables.");
            }

            return false;
        }
#endif
    }
}