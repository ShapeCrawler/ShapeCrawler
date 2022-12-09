using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing.ShapeFill;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal class SCTableCell : ITableCell, ITextFrameContainer
{
    private readonly ResettableLazy<TextFrame> textFrame;
    private readonly ResettableLazy<ShapeFill> fill;

    internal SCTableCell(SCTableRow tableRow, A.TableCell aTableCell, int rowIndex, int columnIndex)
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

    private SCTableRow ParentTableRow { get; }

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