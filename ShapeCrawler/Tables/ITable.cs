using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Collections;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table on a slide.
/// </summary>
public interface ITable : IShape
{
    /// <summary>
    ///     Gets table columns.
    /// </summary>
    IReadOnlyList<IColumn> Columns { get; }

    /// <summary>
    ///     Gets table rows.
    /// </summary>
    IRowCollection Rows { get; }

    /// <summary>
    ///     Gets cell by row and column indexes.
    /// </summary>
    ICell this[int rowIndex, int columnIndex] { get; }

    /// <summary>
    ///     Merge neighbor cells.
    /// </summary>
    void MergeCells(ICell cell1, ICell cell2);
}

internal sealed class SCTable : SCSlideShape, ITable
{
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly ResettableLazy<SCRowCollection> rowCollection;

    internal SCTable(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOrLayout, SCGroupShape groupSCShape)
        : base(pShapeTreeChild, slideOrLayout, groupSCShape)
    {
        this.rowCollection =
            new ResettableLazy<SCRowCollection>(() => SCRowCollection.Create(this, (P.GraphicFrame)this.PShapeTreesChild));
        this.pGraphicFrame = (P.GraphicFrame)pShapeTreeChild;
    }
    
    internal SCTable(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOrLayout)
        : base(pShapeTreeChild, slideOrLayout)
    {
        this.rowCollection =
            new ResettableLazy<SCRowCollection>(() => SCRowCollection.Create(this, (P.GraphicFrame)this.PShapeTreesChild));
        this.pGraphicFrame = (P.GraphicFrame)pShapeTreeChild;
    }

    public override SCShapeType ShapeType => SCShapeType.Table;

    public IReadOnlyList<IColumn> Columns => this.GetColumnList(); // TODO: make lazy

    public IRowCollection Rows => this.rowCollection.Value;

    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    private A.Table ATable => this.pGraphicFrame.GetATable();

    public ICell this[int rowIndex, int columnIndex] => this.Rows[rowIndex].Cells[columnIndex];

    public void MergeCells(ICell inputCell1, ICell inputCell2) // TODO: Optimize method
    {
        SCCell cell1 = (SCCell)inputCell1;
        SCCell cell2 = (SCCell)inputCell2;
        if (CannotBeMerged(cell1, cell2))
        {
            return;
        }

        int minRowIndex = cell1.RowIndex < cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
        int maxRowIndex = cell1.RowIndex > cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
        int minColIndex = cell1.ColumnIndex < cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;
        int maxColIndex = cell1.ColumnIndex > cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;

        var aTableRows = this.ATable.Elements<A.TableRow>().ToList();
        if (minColIndex != maxColIndex)
        {
            this.MergeHorizontal(maxColIndex, minColIndex, minRowIndex, maxRowIndex, aTableRows);
        }

        // Vertical merging
        if (minRowIndex != maxRowIndex)
        {
            // Set row span value for the first cell in the merged cells
            var verticalMergingCount = maxRowIndex - minRowIndex + 1;
            var rowSpanCells = aTableRows[minRowIndex].Elements<A.TableCell>()
                .Skip(minColIndex)
                .Take(maxColIndex + 1);
            foreach (var aTblCell in rowSpanCells)
            {
                aTblCell.RowSpan = new Int32Value(verticalMergingCount);
            }

            // Set vertical merging flag
            foreach (var aTableRow in aTableRows.Skip(minRowIndex + 1).Take(maxRowIndex))
            {
                foreach (A.TableCell aTblCell in aTableRow.Elements<A.TableCell>().Take(maxColIndex + 1))
                {
                    aTblCell.VerticalMerge = new BooleanValue(true);
                    this.MergeParagraphs(minRowIndex, minColIndex, aTblCell);
                }
            }
        }

        // Delete a:gridCol and a:tc elements if all columns are merged
        for (int colIdx = 0; colIdx < this.Columns.Count;)
        {
            int? gridSpan = ((SCCell)this.Rows[0].Cells[colIdx]).ATableCell.GridSpan?.Value;
            if (gridSpan > 1 && this.Rows.All(row =>
                    ((SCCell)row.Cells[colIdx]).ATableCell.GridSpan?.Value == gridSpan))
            {
                int deleteColumnCount = gridSpan.Value - 1;

                // Delete a:gridCol elements
                foreach (SCColumn column in this.Columns.Skip(colIdx + 1).Take(deleteColumnCount))
                {
                    column.AGridColumn.Remove();
                    this.Columns[colIdx].Width += column.Width; // append width of deleting column to merged column
                }

                // Delete a:tc elements
                foreach (A.TableRow aTblRow in aTableRows)
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

        // Delete a:tr if need
        for (var rowIdx = 0; rowIdx < this.Rows.Count;)
        {
            var rowCells = this.Rows[rowIdx].Cells.OfType<SCCell>().ToList();
            var firstRowCell = rowCells[0];
            var rowSpan = firstRowCell.ATableCell.RowSpan?.Value;
            if (rowSpan > 1 && rowCells.All(cell => cell.ATableCell.RowSpan?.Value == rowSpan))
            {
                int deleteRowsCount = rowSpan.Value - 1;

                // Delete a:gridCol elements
                foreach (var row in this.Rows.Skip(rowIdx + 1).Take(deleteRowsCount))
                {
                    ((SCRow)row).ATableRow.Remove();
                    this.Rows[rowIdx].Height += row.Height;
                }

                rowIdx += rowSpan.Value;
                continue;
            }

            rowIdx++;
        }

        this.rowCollection.Reset();
    }

    internal override void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal IRow AppendRow(A.TableRow row)
    {
        this.ATable.AppendChild(row);

        // reset row collection so this.Rows will include the recently added row
        this.rowCollection.Reset();

        // the new row is the last one in the row collection
        return this.Rows.Last();
    }
    
    protected override void SetXCoordinate(int xPx)
    {
        var pXfrm = this.pGraphicFrame.Transform;
        if (pXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value;
            var xEmu = UnitConverter.HorizontalPixelToEmu(xPx);
            var yEmu = UnitConverter.HorizontalPixelToEmu(referencedShape!.Y);
            var wEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            
            this.pGraphicFrame.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            pXfrm.Offset!.X = UnitConverter.HorizontalPixelToEmu(xPx);
        }
    }

    protected override void SetYCoordinate(int yPx)
    {
        var pXfrm = this.pGraphicFrame.Transform;
        if (pXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value!;
            var xEmu = UnitConverter.HorizontalPixelToEmu(referencedShape.X);
            var yEmu = UnitConverter.HorizontalPixelToEmu(yPx);
            var wEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            this.pGraphicFrame.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            pXfrm.Offset!.Y = UnitConverter.HorizontalPixelToEmu(yPx);
        }
    }

    protected override void SetWidth(int wPx)
    {
        var pXfrm = this.pGraphicFrame.Transform;
        if (pXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value;
            var xEmu = UnitConverter.HorizontalPixelToEmu(referencedShape!.X);
            var yEmu = UnitConverter.HorizontalPixelToEmu(referencedShape.Y);
            var wEmu = UnitConverter.VerticalPixelToEmu(wPx);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            this.pGraphicFrame.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            pXfrm.Extents!.Cx = UnitConverter.HorizontalPixelToEmu(wPx);
        }
    }
    
    protected override void SetHeight(int hPx)
    {
        var pXfrm = this.pGraphicFrame.Transform;
        if (pXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value;
            var xEmu = UnitConverter.HorizontalPixelToEmu(referencedShape!.X);
            var yEmu = UnitConverter.HorizontalPixelToEmu(referencedShape.Y);
            var wEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(hPx);
            this.pGraphicFrame.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            pXfrm.Extents!.Cy = UnitConverter.HorizontalPixelToEmu(hPx);
        }
    }

    private static bool CannotBeMerged(SCCell cell1, SCCell cell2)
    {
        if (cell1 == cell2)
        {
            // The cells are already merged
            return true;
        }

        return false;
    }
    
    private void MergeParagraphs(int minRowIndex, int minColIndex, A.TableCell aTblCell)
    {
        A.TextBody? mergedCellTextBody = ((SCCell)this[minRowIndex, minColIndex]).ATableCell.TextBody;
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
    
    private void MergeHorizontal(int maxColIndex, int minColIndex, int minRowIndex, int maxRowIndex, List<A.TableRow> aTableRows)
    {
        int horizontalMergingCount = maxColIndex - minColIndex + 1;
        for (int rowIdx = minRowIndex; rowIdx <= maxRowIndex; rowIdx++)
        {
            A.TableCell[] rowATblCells = aTableRows[rowIdx].Elements<A.TableCell>().ToArray();
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
    
    private IReadOnlyList<SCColumn> GetColumnList()
    {
        IEnumerable<A.GridColumn> aGridColumns = this.ATable.TableGrid!.Elements<A.GridColumn>();
        var columnList = new List<SCColumn>(aGridColumns.Count());
        columnList.AddRange(aGridColumns.Select(aGridColumn => new SCColumn(aGridColumn)));

        return columnList;
    }
}