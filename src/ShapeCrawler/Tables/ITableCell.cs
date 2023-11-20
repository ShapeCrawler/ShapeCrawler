using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Texts;
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

internal sealed class TableCell : ITableCell
{
    internal TableCell(TypedOpenXmlPart sdkTypedOpenXmlPart, A.TableCell aTableCell, int rowIndex, int columnIndex)
    {
        this.ATableCell = aTableCell;
        this.RowIndex = rowIndex;
        this.ColumnIndex = columnIndex;
        this.TextFrame = new TextFrame(sdkTypedOpenXmlPart, this.ATableCell.TextBody!);
        var aTcPr = aTableCell.TableCellProperties!;
        this.Fill = new TableCellFill(sdkTypedOpenXmlPart, aTcPr);
    }

    public bool IsMergedCell => this.ATableCell.GridSpan is not null ||
                                this.ATableCell.RowSpan is not null ||
                                this.ATableCell.HorizontalMerge is not null ||
                                this.ATableCell.VerticalMerge is not null;

    public IShapeFill Fill { get; }

    public ITextFrame TextFrame { get; }

    internal A.TableCell ATableCell { get; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }
}