using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a table on a slide.
/// </summary>
public interface ITable : IShape
{
    /// <summary>
    ///     Gets table columns.
    /// </summary>
    ITableColumnCollection Columns { get; }

    /// <summary>
    ///     Gets table rows.
    /// </summary>
    ITableRowCollection Rows { get; }

    /// <summary>
    ///     Gets or sets the table style.
    /// </summary>
    ITableStyle TableStyle { get; set; }

    /// <summary>
    ///    Gets the table style options.
    /// </summary>
    ITableStyleOptions StyleOptions { get; }

    /// <summary>
    ///     Gets cell by row and column indexes.
    /// </summary>
    ITableCell this[int rowIndex, int columnIndex] { get; }

    /// <summary>
    ///     Merge neighbor cells.
    /// </summary>
    void MergeCells(ITableCell cell1, ITableCell cell2);

    /// <summary>
    ///     Updates table fill.
    /// </summary>
    void UpdateFill(string colorHex);

    /// <summary>
    ///     Gets table cell by row and column numbers.
    /// </summary>
    ITableCell Cell(int rowNumber, int columnNumber);
}

internal sealed class Table(
    Shape shape,
    TableRowCollection rows,
    TableColumnCollection columns,
    TableStyleOptions styleOptions,
    P.GraphicFrame pGraphicFrame) : ITable
{
    private ITableStyle? tableStyle;

    public ShapeContent ShapeContent => ShapeContent.Table;

    public IShapeOutline Outline => shape.Outline;

    public IShapeFill Fill => shape.Fill;

    public ITextBox? TextBox => shape.TextBox;

    public double Rotation => shape.Rotation;

    public string SDKXPath => shape.SDKXPath;

    public OpenXmlElement SDKOpenXmlElement => shape.SDKOpenXmlElement;

    public IPresentation Presentation => shape.Presentation;

    public ITableColumnCollection Columns => columns;

    public ITableRowCollection Rows => rows;

    public ITableStyle TableStyle
    {
        get => this.GetTableStyle();
        set => this.SetTableStyle(value);
    }

    public decimal Width
    {
        get => shape.Width;
        set => shape.Width = value;
    }

    public decimal Height
    {
        get => shape.Height;
        set
        {
            var percentNewHeight = value / shape.Height;

            shape.Height = value;

            foreach (var tableRow in this.Rows)
            {
                var row = (TableRow)tableRow;
                row.SetHeight((int)(row.Height * percentNewHeight));
            }
        }
    }

    public int Id => shape.Id;

    public string Name
    {
        get => shape.Name;
        set => shape.Name = value;
    }

    public string AltText
    {
        get => shape.AltText;
        set => shape.AltText = value;
    }

    public bool Hidden => shape.Hidden;

    public PlaceholderType? PlaceholderType => shape.PlaceholderType;

    public string? CustomData
    {
        get => shape.CustomData;
        set => shape.CustomData = value;
    }

    public ITableStyleOptions StyleOptions => styleOptions;

    public bool Removable => true;

    public Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("Updating geometry is not supported for table.");
    }

    public decimal CornerSize
    {
        get => shape.CornerSize;
        set => shape.CornerSize = value;
    }

    public decimal[] Adjustments
    {
        get => shape.Adjustments;
        set => shape.Adjustments = value;
    }

    public decimal X
    {
        get => shape.X;
        set => shape.X = value;
    }

    public decimal Y
    {
        get => shape.Y;
        set => shape.Y = value;
    }

    private A.Table ATable => pGraphicFrame.GetFirstChild<A.Graphic>() !.GraphicData!.GetFirstChild<A.Table>() !;

    public ITableCell this[int rowIndex, int columnIndex] => this.Rows[rowIndex].Cells[columnIndex];

    public void UpdateFill(string colorHex)
    {
        throw new NotImplementedException();
    }

    public ITableCell Cell(int rowNumber, int columnNumber) => this.Rows[rowNumber - 1].Cells[columnNumber - 1];

    public void MergeCells(ITableCell cell1, ITableCell cell2)
    {
        var cell1Internal = (TableCell)cell1;
        var cell2Internal = (TableCell)cell2;
        if (cell1Internal == cell2Internal)
        {
            throw new SCException("Cannot merge the same cells.");
        }

        var minRowIndex = cell1Internal.RowIndex < cell2Internal.RowIndex
            ? cell1Internal.RowIndex
            : cell2Internal.RowIndex;
        var maxRowIndex = cell1Internal.RowIndex > cell2Internal.RowIndex
            ? cell1Internal.RowIndex
            : cell2Internal.RowIndex;
        var minColIndex = cell1Internal.ColumnIndex < cell2Internal.ColumnIndex
            ? cell1Internal.ColumnIndex
            : cell2Internal.ColumnIndex;
        var maxColIndex = cell1Internal.ColumnIndex > cell2Internal.ColumnIndex
            ? cell1Internal.ColumnIndex
            : cell2Internal.ColumnIndex;

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
        this.RemoveRowOnDemand();
    }

    public void Remove() => pGraphicFrame.Remove();

    public ITable AsTable() => this;

    public IMediaShape AsMedia() => shape.AsMedia();

    public void Duplicate() => shape.Duplicate();

    public void SetText(string text) => shape.SetText(text);

    public void SetImage(string imagePath) => shape.SetImage(imagePath);

    public void SetFontName(string fontName) => shape.SetFontName(fontName);

    public void SetFontSize(decimal fontSize) => shape.SetFontSize(fontSize);

    public void SetFontColor(string colorHex) => shape.SetFontColor(colorHex);

    internal void SetTableHeight(decimal value) => shape.Height = value;
    
    internal void CopyTo(P.ShapeTree pShapeTree) => shape.CopyTo(pShapeTree);

    private static void DeleteTableCells(int colIdx, int deleteColumnCount, List<A.TableRow> aTableRows)
    {
        foreach (var aTblRow in aTableRows)
        {
            var removeCells = aTblRow.Elements<A.TableCell>().Skip(colIdx + 1).Take(deleteColumnCount).ToList();
            foreach (var aTblCell in removeCells)
            {
                aTblCell.Remove();
            }
        }
    }

    private void SetTableStyle(ITableStyle style)
    {
        this.ATable.TableProperties!.GetFirstChild<A.TableStyleId>() !.Text = ((TableStyle)style).Guid;
        this.tableStyle = style;
    }

    private ITableStyle GetTableStyle()
    {
        if (this.tableStyle is null)
        {
            var aTableStyleId = this.ATable.TableProperties!.GetFirstChild<A.TableStyleId>() !.Text;
            var style = CommonTableStyles.GetTableStyleByGuid(aTableStyleId) !;
            this.tableStyle = style;
        }

        return this.tableStyle;
    }

    private void RemoveRowOnDemand()
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
            A.TableCell[] rowATblCells = [.. aTableRows[rowIdx].Elements<A.TableCell>()];
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
                colIdx += this.ProcessColumnsWithSameGridSpan(colIdx, topColumnCellSpan.Value, aTableRows);
            }
            else
            {
                colIdx++;
            }
        }
    }

    private int ProcessColumnsWithSameGridSpan(int colIdx, int topColumnCellSpan, List<A.TableRow> aTableRows)
    {
        var deleteColumnCount = topColumnCellSpan - 1;

        // Delete a:gridCol elements and append width of deleting column to merged column
        this.DeleteAndUpdateGridColumns(colIdx, deleteColumnCount);

        // Delete a:tc elements
        DeleteTableCells(colIdx, deleteColumnCount, aTableRows);

        return topColumnCellSpan;
    }

    private void DeleteAndUpdateGridColumns(int colIdx, int deleteColumnCount)
    {
        for (int i = 0; i < deleteColumnCount; i++)
        {
            var column = (Column)this.Columns[colIdx + 1 + i];
            column.AGridColumn.Remove();
            this.Columns[colIdx].Width += column.Width;
        }
    }
}