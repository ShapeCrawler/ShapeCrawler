using System;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table cell.
/// </summary>
public interface ICell
{
    /// <summary>
    ///     Gets text box.
    /// </summary>
    ITextFrame TextFrame { get; }

    /// <summary>
    ///     Gets a value indicating whether cell belongs to merged cell.
    /// </summary>
    bool IsMergedCell { get; }

    /// <summary>
    ///     Gets Shape Fill of the cell.
    /// </summary>
    IShapeFill Fill { get; }
}

internal sealed class SCCell : ICell
{
    private readonly Lazy<SCTextFrame> textFrame;
    private readonly Lazy<SCShapeFill> shapeFill;

    internal SCCell(
        SCRow parentTableRow,
        A.TableCell aTableCell,
        int rowIndex,
        int columnIndex)
    {
        this.ParentTableRow = parentTableRow;
        this.ATableCell = aTableCell;
        this.RowIndex = rowIndex;
        this.ColumnIndex = columnIndex;
        this.textFrame = new Lazy<SCTextFrame>(this.CreateTextFrame);
        var tableCellProperties = aTableCell.TableCellProperties!;
        this.shapeFill = new Lazy<SCShapeFill>(() =>
            new SCCellFill(tableCellProperties));
    }

    public bool IsMergedCell => this.DefineWhetherCellIsMerged();

    public IShapeFill Fill => this.shapeFill.Value;

    public ITextFrame TextFrame => this.textFrame.Value;

    internal A.TableCell ATableCell { get; init; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }

    private SCRow ParentTableRow { get; }

    private SCTextFrame CreateTextFrame()
    {
        return new SCTextFrame(this, this.ATableCell.TextBody!, this.slideStructure, this);
    }

    private bool DefineWhetherCellIsMerged()
    {
        return this.ATableCell.GridSpan is not null ||
               this.ATableCell.RowSpan is not null ||
               this.ATableCell.HorizontalMerge is not null ||
               this.ATableCell.VerticalMerge is not null;
    }
}