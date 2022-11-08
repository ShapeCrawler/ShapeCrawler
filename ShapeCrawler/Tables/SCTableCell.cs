using ShapeCrawler.AutoShapes;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class SCTableCell : ITableCell, ITextFrameContainer
{
    private readonly ResettableLazy<TextFrame> textFrame;

    internal SCTableCell(SCTableRow tableRow, A.TableCell aTableCell, int rowIndex, int columnIndex)
    {
        this.ParentTableRow = tableRow;
        this.ATableCell = aTableCell;
        this.RowIndex = rowIndex;
        this.ColumnIndex = columnIndex;
        this.textFrame = new ResettableLazy<TextFrame>(this.GetTextFrame);
    }

    public bool IsMergedCell => this.DefineWhetherCellIsMerged();

    public Shape Shape => this.ParentTableRow.ParentTable;

    public ITextFrame TextFrame => this.textFrame.Value;

    internal A.TableCell ATableCell { get; init; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }

    private SCTableRow ParentTableRow { get; }

    public void ThrowIfRemoved()
    {
        this.ParentTableRow.ThrowIfRemoved();
    }

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