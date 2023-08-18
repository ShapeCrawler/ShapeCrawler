using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table cell.
/// </summary>
public interface ITableCell
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

internal sealed record SCTableCell : ITableCell
{
    private readonly Lazy<SCTextFrame> textFrame;
    private readonly Lazy<SCCellFill> shapeFill;
    private readonly SCTableRow _parentTableRow;

    internal SCTableCell(
        SCTableRow _parentTableRow,
        A.TableCell aTableCell,
        int rowIndex,
        int columnIndex)
    {
        this._parentTableRow = _parentTableRow;
        this.ATableCell = aTableCell;
        this.RowIndex = rowIndex;
        this.ColumnIndex = columnIndex;
        this.textFrame = new Lazy<SCTextFrame>(this.CreateTextFrame);
        var tableCellProperties = aTableCell.TableCellProperties!;
        this.shapeFill = new Lazy<SCCellFill>(() =>
            new SCCellFill(tableCellProperties, this));
    }

    public bool IsMergedCell => this.DefineWhetherCellIsMerged();

    public IShapeFill Fill => this.shapeFill.Value;

    public ITextFrame TextFrame => this.textFrame.Value;

    internal A.TableCell ATableCell { get; init; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }

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

    internal SlidePart SDKSlidePart()
    {
        return this._parentTableRow.SDKSlidePart();
    }
}