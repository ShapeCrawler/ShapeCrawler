using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
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

    /// <summary>
    ///     Removes a column at specified index.
    /// </summary>
    void RemoveColumnAt(int columnIndex);
}

internal sealed class SCTable : SCShape, ITable
{
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly ResetAbleLazy<SCRowCollection> rowCollection;

    internal SCTable(
        OpenXmlCompositeElement pShapeTreeChild, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pShapeTreeChild, parentSlideObject, parentShapeCollection)
    {
        this.rowCollection =
            new ResetAbleLazy<SCRowCollection>(() => SCRowCollection.Create(this, (P.GraphicFrame)this.PShapeTreeChild));
        this.pGraphicFrame = (P.GraphicFrame)pShapeTreeChild;
    }
    
    public override SCShapeType ShapeType => SCShapeType.Table;

    public IReadOnlyList<IColumn> Columns => this.GetColumnList(); // TODO: make lazy

    public IRowCollection Rows => this.rowCollection.Value;

    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    private A.Table ATable => this.pGraphicFrame.GetATable();

    public ICell this[int rowIndex, int columnIndex] => this.Rows[rowIndex].Cells[columnIndex];

    public void RemoveColumnAt(int columnIndex)
    {
        var column = (SCColumn)this.Columns[columnIndex];
        column.AGridColumn.Remove();
        
        var aTableRows = this.ATable.Elements<A.TableRow>();

        foreach (var aTableRow in aTableRows)
        {
            var aTableCell = aTableRow.Elements<A.TableCell>().ElementAt(columnIndex);
            aTableCell.Remove();
        }
    }

    public void MergeCells(ICell inputCell1, ICell inputCell2)
    {
        var cell1 = (SCCell)inputCell1;
        var cell2 = (SCCell)inputCell2;
        if (cell1 == cell2)   
        {
            throw new SCException("Cannot merge the same cells.");
        }

        var minRowIndex = cell1.RowIndex < cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
        var maxRowIndex = cell1.RowIndex > cell2.RowIndex ? cell1.RowIndex : cell2.RowIndex;
        var minColIndex = cell1.ColumnIndex < cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;
        var maxColIndex = cell1.ColumnIndex > cell2.ColumnIndex ? cell1.ColumnIndex : cell2.ColumnIndex;

        var aTableRows = this.ATable.Elements<A.TableRow>().ToList();
        if (minColIndex != maxColIndex)
        {
            this.MergeHorizontal(maxColIndex, minColIndex, minRowIndex, maxRowIndex, aTableRows);
        }

        if (minRowIndex != maxRowIndex)
        {
            this.MergeVertically(maxRowIndex, minRowIndex, aTableRows, minColIndex, maxColIndex);
        }

        this.RemoveColumnIfNeeded(aTableRows);

        // Delete a:tr if needed
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

    private void RemoveColumnIfNeeded(List<A.TableRow> aTableRows)
    {
        // Delete a:gridCol and a:tc elements if all columns are merged
        for (var colIdx = 0; colIdx < this.Columns.Count;)
        {
            var topColumnCell = ((SCRow)this.Rows[0]).ATableRow.Elements<A.TableCell>().ToList()[colIdx];
            var topColumnCellSpan = topColumnCell.GridSpan?.Value;
            var nextBottomColumnCells = this.Rows.Select(row => ((SCRow)row).ATableRow.Elements<A.TableCell>().ToList()[colIdx]).ToList();
            var sameGridSpan = nextBottomColumnCells.All(c => c.GridSpan?.Value == topColumnCellSpan);
            if (topColumnCellSpan > 1 && sameGridSpan)
            {
                var deleteColumnCount = topColumnCellSpan.Value - 1;

                // Delete a:gridCol elements and append width of deleting column to merged column
                for (int i = 0; i < deleteColumnCount; i++)
                {
                    var column = (SCColumn)this.Columns[colIdx + 1 + i];
                    column.AGridColumn.Remove();
                    this.Columns[colIdx].Width += column.Width;
                }

                // Delete a:tc elements
                foreach (var aTblRow in aTableRows)
                {
                    var removeCells = aTblRow.Elements<A.TableCell>().Skip(colIdx).Take(deleteColumnCount).ToList();
                    foreach (var aTblCell in removeCells)
                    {
                        aTblCell.Remove();
                    }
                }

                colIdx += topColumnCellSpan.Value;
            }
            else
            {
                colIdx++;
            }
        }
    }

    internal override void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal override IHtmlElement ToHtmlElement()
    {
        throw new NotImplementedException();
    }

    internal override string ToJson()
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
    
    private void MergeVertically(int bottomIndex, int topRowIndex, List<A.TableRow> aTableRows, int leftColIndex, int rightColIndex)
    {
        int verticalMergingCount = bottomIndex - topRowIndex + 1;
    
        // Set row span value for the first cell in the merged cells
        foreach (var aTblCell in aTableRows[topRowIndex].Elements<A.TableCell>().Skip(leftColIndex).Take(rightColIndex + 1))
        {
            aTblCell.RowSpan = new Int32Value(verticalMergingCount);
        }

        // Set vertical merging flag
        foreach (var aTableRow in aTableRows.Skip(topRowIndex + 1).Take(bottomIndex - topRowIndex))
        {
            foreach (var aTc in aTableRow.Elements<A.TableCell>().Skip(leftColIndex).Take(rightColIndex - leftColIndex + 1))
            {
                aTc.VerticalMerge = new BooleanValue(true);
                this.MergeParagraphs(topRowIndex, leftColIndex, aTc);
            }
        }
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