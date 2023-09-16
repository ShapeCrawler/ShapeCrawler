using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
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
    ITableCell this[int rowIndex, int columnIndex] { get; }

    /// <summary>
    ///     Merge neighbor cells.
    /// </summary>
    void MergeCells(ITableCell cell1, ITableCell cell2);

    /// <summary>
    ///     Removes a column at specified index.
    /// </summary>
    void RemoveColumnAt(int columnIndex);

    void UpdateFill(string colorHex);
}

internal sealed class SlideTable : CopyableShape, ITable, IRemoveable 
{
    private readonly SlidePart sdkSlidePart;
    private readonly P.GraphicFrame pGraphicFrame;
    private readonly ResetableLazy<SlideTableRows> rowCollection;

    internal SlideTable(SlidePart sdkSlidePart, OpenXmlCompositeElement pShapeTreeElement)
        : base(pShapeTreeElement)
    {
        this.sdkSlidePart = sdkSlidePart;
        var graphicFrame = (P.GraphicFrame)pShapeTreeElement;
        this.rowCollection =
            new ResetableLazy<SlideTableRows>(() => new SlideTableRows(this.sdkSlidePart, graphicFrame));
        this.pGraphicFrame = (P.GraphicFrame)pShapeTreeElement;
    }

    public override SCShapeType ShapeType => SCShapeType.Table;
    public IReadOnlyList<IColumn> Columns => this.GetColumnList(); // TODO: make lazy
    public IRowCollection Rows => this.rowCollection.Value;
    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    private A.Table ATable => this.pGraphicFrame.GetATable();

    public ITableCell this[int rowIndex, int columnIndex] => this.Rows[rowIndex].Cells[columnIndex];

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

    public void UpdateFill(string colorHex)
    {
        throw new NotImplementedException();
    }

    public void MergeCells(ITableCell inputCell1, ITableCell inputCell2)
    {
        var cell1 = (TableCell)inputCell1;
        var cell2 = (TableCell)inputCell2;
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
        this.RemoveRowIfNeeded();

        this.rowCollection.Reset();
    }

    internal void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new NotImplementedException();
    }

    internal string ToJson()
    {
        throw new NotImplementedException();
    }

    private static bool CannotBeMerged(TableCell cell1, TableCell cell2)
    {
        if (cell1 == cell2)
        {
            // The cells are already merged
            return true;
        }

        return false;
    }

    private void RemoveRowIfNeeded()
    {
        // Delete a:tr if needed
        for (var rowIdx = 0; rowIdx < this.Rows.Count;)
        {
            var allRowCells = this.Rows[rowIdx].Cells.OfType<TableCell>().ToList();
            var firstRowCell = allRowCells[0];
            var firstRowCellSpan = firstRowCell.ATableCell.RowSpan?.Value;
            if (firstRowCellSpan > 1 && allRowCells.All(cell => cell.ATableCell.RowSpan?.Value == firstRowCellSpan))
            {
                int deleteRowsCount = firstRowCellSpan.Value - 1;

                foreach (var row in this.Rows.Skip(rowIdx + 1).Take(deleteRowsCount))
                {
                    ((SlideTableRow)row).ATableRow.Remove();
                    this.Rows[rowIdx].Height += row.Height;
                }

                rowIdx += firstRowCellSpan.Value;
                continue;
            }

            rowIdx++;
        }
    }

    private void MergeVertically(
        int bottomIndex,
        int topRowIndex,
        List<A.TableRow> aTableRows,
        int leftColIndex,
        int rightColIndex)
    {
        int verticalMergingCount = bottomIndex - topRowIndex + 1;
        var numMergingCells = rightColIndex - leftColIndex + 1;
        var horizontalCells =
            aTableRows[topRowIndex].Elements<A.TableCell>().Skip(leftColIndex).Take(numMergingCells);
        foreach (var aTblCell in horizontalCells)
        {
            aTblCell.RowSpan = new Int32Value(verticalMergingCount);
        }

        // Set vertical merging flag
        foreach (var aTableRow in aTableRows.Skip(topRowIndex + 1).Take(bottomIndex - topRowIndex))
        {
            foreach (var aTc in aTableRow.Elements<A.TableCell>().Skip(leftColIndex)
                         .Take(rightColIndex - leftColIndex + 1))
            {
                aTc.VerticalMerge = new BooleanValue(true);
                this.MergeParagraphs(topRowIndex, leftColIndex, aTc);
            }
        }
    }

    private void MergeParagraphs(int minRowIndex, int minColIndex, A.TableCell aTblCell)
    {
        A.TextBody? mergedCellTextBody = ((TableCell)this[minRowIndex, minColIndex]).ATableCell.TextBody;
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

    private void MergeHorizontal(int maxColIndex, int minColIndex, int minRowIndex, int maxRowIndex,
        List<A.TableRow> aTableRows)
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

    private void RemoveColumnIfNeeded(List<A.TableRow> aTableRows)
    {
        // Delete a:gridCol and a:tc elements if all columns are merged
        for (var colIdx = 0; colIdx < this.Columns.Count;)
        {
            var topColumnCell = ((SlideTableRow)this.Rows[0]).ATableRow.Elements<A.TableCell>().ToList()[colIdx];
            var topColumnCellSpan = topColumnCell.GridSpan?.Value;
            var nextBottomColumnCells = this.Rows
                .Select(row => ((SlideTableRow)row).ATableRow.Elements<A.TableCell>().ToList()[colIdx]).ToList();
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

    public void Remove()
    {
        throw new NotImplementedException();
    }
}