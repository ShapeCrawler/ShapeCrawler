using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Tables;
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
    ITableRows Rows { get; }

    /// <summary>
    ///     Gets or sets the table style.
    /// </summary>
    ITableStyle TableStyle { get; set; }

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

    /// <summary>
    ///     Updates table fill.
    /// </summary>
    void UpdateFill(string colorHex);
}

internal sealed class Table : CopyableShape, ITable
{
    private readonly P.GraphicFrame pGraphicFrame;
    private ITableStyle? tableStyle;

    internal Table(OpenXmlPart sdkTypedOpenXmlPart, OpenXmlCompositeElement pShapeTreeElement)
        : base(sdkTypedOpenXmlPart, pShapeTreeElement)
    {
        this.pGraphicFrame = (P.GraphicFrame)pShapeTreeElement;
        this.Rows = new TableRows(sdkTypedOpenXmlPart, this.pGraphicFrame);
    }

    public override ShapeType ShapeType => ShapeType.Table;

    public IReadOnlyList<IColumn> Columns => this.GetColumnList(); // TODO: make lazy

    public ITableRows Rows { get; }

    public ITableStyle TableStyle
    {
        get => this.GetTableStyle();
        set => this.SetTableStyle(value);
    }

    public override bool Removeable => true;

    public override Geometry GeometryType => Geometry.Rectangle;

    private A.Table ATable => this.pGraphicFrame.ATable();

    public ITableCell this[int rowIndex, int columnIndex] => this.Rows[rowIndex].Cells[columnIndex];

    public void RemoveColumnAt(int columnIndex)
    {
        var column = (Column)this.Columns[columnIndex];
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

    public void MergeCells(ITableCell cell1, ITableCell cell2)
    {
        var cell1Internal = (TableCell)cell1;
        var cell2Internal = (TableCell)cell2;
        if (cell1Internal == cell2Internal)
        {
            throw new SCException("Cannot merge the same cells.");
        }

        var minRowIndex = cell1Internal.RowIndex < cell2Internal.RowIndex ? cell1Internal.RowIndex : cell2Internal.RowIndex;
        var maxRowIndex = cell1Internal.RowIndex > cell2Internal.RowIndex ? cell1Internal.RowIndex : cell2Internal.RowIndex;
        var minColIndex = cell1Internal.ColumnIndex < cell2Internal.ColumnIndex ? cell1Internal.ColumnIndex : cell2Internal.ColumnIndex;
        var maxColIndex = cell1Internal.ColumnIndex > cell2Internal.ColumnIndex ? cell1Internal.ColumnIndex : cell2Internal.ColumnIndex;

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
    }

    public override void Remove() => this.pGraphicFrame.Remove();

    public override ITable AsTable() => this;

    internal void Draw(SKCanvas canvas)
    {
        throw new NotImplementedException();
    }

    internal string ToJson()
    {
        throw new NotImplementedException();
    }
    
    private void SetTableStyle(ITableStyle style)
    {
        this.ATable.TableProperties!.GetFirstChild<A.TableStyleId>() !.Text = style.Guid;
        this.tableStyle = style;
    }

    private ITableStyle GetTableStyle()
    {
        if (this.tableStyle is null)
        {
            var tableStyleId = this.ATable.TableProperties!.GetFirstChild<A.TableStyleId>() !.Text;

            var style = CommonTableStyles.GetTableStyleByGuid(tableStyleId) !;

            // style ??= new TableStyle("Custom Style", tableStyleId);
            this.tableStyle = style;
        }

        return this.tableStyle;
    }

    private void RemoveRowIfNeeded()
    {
        int rowIdx = 0;

        while (rowIdx < this.Rows.Count)
        {
            var cells = this.Rows[rowIdx].Cells.OfType<TableCell>().ToList();
            var firstCell = cells[0];
            var firstCellSpan = firstCell.ATableCell.RowSpan?.Value;

            if (firstCellSpan > 1 && cells.All(cell => cell.ATableCell.RowSpan?.Value == firstCellSpan))
            {
                int deleteRowsCount = firstCellSpan.Value - 1;
                foreach (var row in this.Rows.Skip(rowIdx + 1).Take(deleteRowsCount))
                {
                    ((TableRow)row).ATableRow.Remove();
                    this.Rows[rowIdx].Height += row.Height;
                }

                rowIdx += firstCellSpan.Value;
            }
            else
            {
                rowIdx++;
            }
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

    private void MergeHorizontal(
        int maxColIndex,
        int minColIndex,
        int minRowIndex,
        int maxRowIndex,
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

    private IReadOnlyList<Column> GetColumnList()
    {
        IEnumerable<A.GridColumn> aGridColumns = this.ATable.TableGrid!.Elements<A.GridColumn>();
        var columnList = new List<Column>(aGridColumns.Count());
        columnList.AddRange(aGridColumns.Select(aGridColumn => new Column(aGridColumn)));

        return columnList;
    }

    private void RemoveColumnIfNeeded(List<A.TableRow> aTableRows)
    {
        int colIdx = 0;
        while (colIdx < this.Columns.Count)
        {
            var topColumnCell = ((TableRow)this.Rows[0]).ATableRow.Elements<A.TableCell>().ToList()[colIdx];
            var topColumnCellSpan = topColumnCell.GridSpan?.Value;
            var nextBottomColumnCells = this.Rows
                .Select(row => ((TableRow)row).ATableRow.Elements<A.TableCell>().ToList()[colIdx]).ToList();
            var sameGridSpan = nextBottomColumnCells.All(c => c.GridSpan?.Value == topColumnCellSpan);
            if (topColumnCellSpan > 1 && sameGridSpan)
            {
                var deleteColumnCount = topColumnCellSpan.Value - 1;

                // Delete a:gridCol elements and append width of deleting column to merged column
                for (int i = 0; i < deleteColumnCount; i++)
                {
                    var column = (Column)this.Columns[colIdx + 1 + i];
                    column.AGridColumn.Remove();
                    this.Columns[colIdx].Width += column.Width;
                }

                // Delete a:tc elements
                foreach (var aTblRow in aTableRows)
                {
                    var removeCells = aTblRow.Elements<A.TableCell>().Skip(colIdx + 1).Take(deleteColumnCount).ToList();
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
}