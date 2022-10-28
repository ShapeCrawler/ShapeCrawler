using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OneOf;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Tables;

internal class SlideTable : SlideShape, ITable
{
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly ResettableLazy<RowCollection> rowCollection;

    internal SlideTable(OpenXmlCompositeElement childOfPShapeTrees, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOrLayout, SCGroupShape groupShape)
        : base(childOfPShapeTrees, slideOrLayout, groupShape)
    {
        this.rowCollection =
            new ResettableLazy<RowCollection>(() => RowCollection.Create(this, (P.GraphicFrame)this.PShapeTreesChild));
        this.pGraphicFrame = (P.GraphicFrame)childOfPShapeTrees;
    }

    public override SCShapeType ShapeType => SCShapeType.Table;

    public IReadOnlyList<SCColumn> Columns => this.GetColumnList(); // TODO: make lazy

    public IRowCollection Rows => this.rowCollection.Value;

    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    private A.Table ATable => this.pGraphicFrame.GetATable();

    public ITableCell this[int rowIndex, int columnIndex] => this.Rows[rowIndex].Cells[columnIndex];

    public void MergeCells(ITableCell inputCell1, ITableCell inputCell2) // TODO: Optimize method
    {
        SCTableCell cell1 = (SCTableCell)inputCell1;
        SCTableCell cell2 = (SCTableCell)inputCell2;
        if (CannotBeMerged(cell1, cell2))
        {
            return;
        }

        int minRowIndex = cell1.RowIndex < cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
        int maxRowIndex = cell1.RowIndex > cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
        int minColIndex = cell1.ColumnIndex < cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;
        int maxColIndex = cell1.ColumnIndex > cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;

        // Horizontal merging
        List<A.TableRow> aTableRowList = this.ATable.Elements<A.TableRow>().ToList();
        if (minColIndex != maxColIndex)
        {
            int horizontalMergingCount = maxColIndex - minColIndex + 1;
            for (int rowIdx = minRowIndex; rowIdx <= maxRowIndex; rowIdx++)
            {
                A.TableCell[] rowATblCells = aTableRowList[rowIdx].Elements<A.TableCell>().ToArray();
                A.TableCell firstMergingCell = rowATblCells[minColIndex];
                firstMergingCell.GridSpan = new Int32Value(horizontalMergingCount);
                Span<A.TableCell> nextMergingCells =
                    new Span<A.TableCell>(rowATblCells, minColIndex + 1, horizontalMergingCount - 1);
                foreach (A.TableCell aTblCell in nextMergingCells)
                {
                    aTblCell.HorizontalMerge = new BooleanValue(true);

                    this.MergeParagraphs(minRowIndex, minColIndex, aTblCell);
                }
            }
        }

        // Vertical merging
        if (minRowIndex != maxRowIndex)
        {
            // Set row span value for the first cell in the merged cells
            int verticalMergingCount = maxRowIndex - minRowIndex + 1;
            IEnumerable<A.TableCell> rowSpanCells = aTableRowList[minRowIndex].Elements<A.TableCell>()
                .Skip(minColIndex)
                .Take(maxColIndex + 1);
            foreach (A.TableCell aTblCell in rowSpanCells)
            {
                aTblCell.RowSpan = new Int32Value(verticalMergingCount);
            }

            // Set vertical merging flag
            foreach (A.TableRow aTblRow in aTableRowList.Skip(minRowIndex + 1).Take(maxRowIndex))
            {
                foreach (A.TableCell aTblCell in aTblRow.Elements<A.TableCell>().Take(maxColIndex + 1))
                {
                    aTblCell.VerticalMerge = new BooleanValue(true);

                    this.MergeParagraphs(minRowIndex, minColIndex, aTblCell);
                }
            }
        }

        // Delete a:gridCol and a:tc elements if all columns are merged
        for (int colIdx = 0; colIdx < this.Columns.Count;)
        {
            int? gridSpan = ((SCTableCell)this.Rows[0].Cells[colIdx]).ATableCell.GridSpan?.Value;
            if (gridSpan > 1 && this.Rows.All(row =>
                    ((SCTableCell)row.Cells[colIdx]).ATableCell.GridSpan?.Value == gridSpan))
            {
                int deleteColumnCount = gridSpan.Value - 1;

                // Delete a:gridCol elements
                foreach (SCColumn column in this.Columns.Skip(colIdx + 1).Take(deleteColumnCount))
                {
                    column.AGridColumn.Remove();
                    this.Columns[colIdx].Width += column.Width; // append width of deleting column to merged column
                }

                // Delete a:tc elements
                foreach (A.TableRow aTblRow in aTableRowList)
                {
                    IEnumerable<A.TableCell> removeCells =
                        aTblRow.Elements<A.TableCell>().Skip(colIdx).Take(deleteColumnCount);
                    foreach (A.TableCell aTblCell in removeCells)
                    {
                        aTblCell.Remove();
                    }
                }

                colIdx += gridSpan.Value;
                continue;
            }

            colIdx++;
        }

        // Delete a:tr
        for (int rowIdx = 0; rowIdx < this.Rows.Count;)
        {
            int? rowSpan = ((SCTableCell)this.Rows[rowIdx].Cells[0]).ATableCell.RowSpan?.Value;
            if (rowSpan > 1 && this.Rows[rowIdx].Cells.All(c => ((SCTableCell)c).ATableCell.RowSpan?.Value == rowSpan))
            {
                int deleteRowsCount = rowSpan.Value - 1;

                // Delete a:gridCol elements
                foreach (var row in this.Rows.Skip(rowIdx + 1).Take(deleteRowsCount))
                {
                    ((SCTableRow)row).ATableRow.Remove();
                    this.Rows[rowIdx].Height += row.Height;
                }

                rowIdx += rowSpan.Value;
                continue;
            }

            rowIdx++;
        }

        this.rowCollection.Reset();
    }

    internal ITableRow AppendRow(A.TableRow row)
    {
        this.ATable.AppendChild(row);

        // reset row collection so this.Rows will include the recently added row
        this.rowCollection.Reset();

        // the new row is the last one in the row collection
        return this.Rows.Last();
    }

    private void MergeParagraphs(int minRowIndex, int minColIndex, A.TableCell aTblCell)
    {
        A.TextBody? mergedCellTextBody = ((SCTableCell)this[minRowIndex, minColIndex]).ATableCell.TextBody;
        bool hasMoreOnePara = false;
        IEnumerable<A.Paragraph> aParagraphsWithARun =
            aTblCell.TextBody!.Elements<A.Paragraph>().Where(p => !p.IsEmpty());
        foreach (A.Paragraph aParagraph in aParagraphsWithARun)
        {
            mergedCellTextBody!.Append(aParagraph.CloneNode(true));
            hasMoreOnePara = true;
        }

        if (hasMoreOnePara)
        {
            foreach (A.Paragraph aParagraph in mergedCellTextBody!.Elements<A.Paragraph>().Where(p => p.IsEmpty()))
            {
                aParagraph.Remove();
            }
        }
    }

    private IReadOnlyList<SCColumn> GetColumnList()
    {
        IEnumerable<A.GridColumn> aGridColumns = this.ATable.TableGrid!.Elements<A.GridColumn>();
        var columnList = new List<SCColumn>(aGridColumns.Count());
        columnList.AddRange(aGridColumns.Select(aGridColumn => new SCColumn(aGridColumn)));

        return columnList;
    }

    private static bool CannotBeMerged(SCTableCell cell1, SCTableCell cell2)
    {
        if (cell1 == cell2)
        {
            // The cells are already merged
            return true;
        }

        return false;
    }

}