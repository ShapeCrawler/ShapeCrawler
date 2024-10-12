using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tables;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a table cell.
/// </summary>
public interface ITableCell
{
    /// <summary>
    ///     Gets text box.
    /// </summary>
    ITextBox TextBox { get; }

    /// <summary>
    ///     Gets a value indicating whether cell belongs to merged cell.
    /// </summary>
    bool IsMergedCell { get; }

    /// <summary>
    ///     Gets Shape Fill of the cell.
    /// </summary>
    IShapeFill Fill { get; }

    /// <summary>
    ///     Gets the Top Border.
    /// </summary>
    IBorder TopBorder { get; }
    
    /// <summary>
    ///     Gets the Bottom Border.
    /// </summary>
    IBorder BottomBorder { get; }
    
    /// <summary>
    ///     Gets the Left Border.
    /// </summary>
    IBorder LeftBorder { get; }
    
    /// <summary>
    ///     Gets the Right Border.
    /// </summary>
    IBorder RightBorder { get; }
}

internal sealed class TableCell : ITableCell
{
    internal TableCell(OpenXmlPart sdkTypedOpenXmlPart, A.TableCell aTableCell, int rowIndex, int columnIndex)
    {
        this.ATableCell = aTableCell;
        this.RowIndex = rowIndex;
        this.ColumnIndex = columnIndex;
        this.TextBox = new TextBox(sdkTypedOpenXmlPart, this.ATableCell.TextBody!);
        var aTcPr = aTableCell.TableCellProperties!;
        this.Fill = new TableCellFill(sdkTypedOpenXmlPart, aTcPr);
        this.TopBorder = new TopBorder(aTableCell.TableCellProperties!);
        this.BottomBorder = new BottomBorder(aTableCell.TableCellProperties!);
        this.LeftBorder = new LeftBorder(aTableCell.TableCellProperties!);
        this.RightBorder = new RightBorder(aTableCell.TableCellProperties!);
    }

    public bool IsMergedCell => this.ATableCell.GridSpan is not null ||
                                this.ATableCell.RowSpan is not null ||
                                this.ATableCell.HorizontalMerge is not null ||
                                this.ATableCell.VerticalMerge is not null;

    public IShapeFill Fill { get; }

    public IBorder TopBorder { get; }
    
    public IBorder BottomBorder { get; }

    public IBorder LeftBorder { get; }

    public IBorder RightBorder { get; }

    public ITextBox TextBox { get; }

    internal A.TableCell ATableCell { get; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }
}