using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing.ShapeFill;
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

internal class SCCell : ICell, ITextFrameContainer
{
    private readonly ResettableLazy<TextFrame> textFrame;
    private readonly ResettableLazy<ShapeFill> fill;

    internal SCCell(SCRow tableRow, A.TableCell aTableCell, int rowIndex, int columnIndex)
    {
        this.ParentTableRow = tableRow;
        this.ATableCell = aTableCell;
        this.RowIndex = rowIndex;
        this.ColumnIndex = columnIndex;
        this.textFrame = new ResettableLazy<TextFrame>(this.GetTextFrame);
        var slideObject = tableRow.ParentTable.SlideObject;
        var framePr = aTableCell.TableCellProperties!;
        this.fill = new ResettableLazy<ShapeFill>(() => new CellFill((SlideObject)slideObject, framePr));
    }

    public bool IsMergedCell => this.DefineWhetherCellIsMerged();

    public IShapeFill Fill => this.fill.Value;

    public Shape Shape => this.ParentTableRow.ParentTable;

    public ITextFrame TextFrame => this.textFrame.Value;

    internal A.TableCell ATableCell { get; init; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }

    private SCRow ParentTableRow { get; }

    private TextFrame GetTextFrame()
    {
        return new TextFrame(this, this.ATableCell.TextBody!);
    }

    private bool DefineWhetherCellIsMerged()
    {
        return this.ATableCell.GridSpan is not null ||
               this.ATableCell.RowSpan is not null ||
               this.ATableCell.HorizontalMerge is not null ||
               this.ATableCell.VerticalMerge is not null;
    }
}