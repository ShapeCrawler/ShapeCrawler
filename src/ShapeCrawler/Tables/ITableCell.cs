using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
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

    /// <summary>
    ///     Gets the top border.
    /// </summary>
    ITopBorder TopBorder { get; }
}

/// <summary>
///     Represents a top border of a table cell.
/// </summary>
public interface ITopBorder
{
    /// <summary>
    ///     Gets or sets border width in points.
    /// </summary>
    float Width { get; set; }
}

internal sealed class TableCell : ITableCell
{
    internal TableCell(OpenXmlPart sdkTypedOpenXmlPart, A.TableCell aTableCell, int rowIndex, int columnIndex)
    {
        this.ATableCell = aTableCell;
        this.RowIndex = rowIndex;
        this.ColumnIndex = columnIndex;
        this.TextFrame = new TextFrame(sdkTypedOpenXmlPart, this.ATableCell.TextBody!);
        var aTcPr = aTableCell.TableCellProperties!;
        this.Fill = new TableCellFill(sdkTypedOpenXmlPart, aTcPr);
        this.TopBorder = new TopBorder(aTableCell.TableCellProperties!);
    }

    public bool IsMergedCell => this.ATableCell.GridSpan is not null ||
                                this.ATableCell.RowSpan is not null ||
                                this.ATableCell.HorizontalMerge is not null ||
                                this.ATableCell.VerticalMerge is not null;

    public IShapeFill Fill { get; }
    public ITopBorder TopBorder { get; }

    public ITextFrame TextFrame { get; }

    internal A.TableCell ATableCell { get; }

    internal int RowIndex { get; }

    internal int ColumnIndex { get; }
}

internal class TopBorder : ITopBorder
{
    private readonly A.TableCellProperties aTableCellProperties;

    internal TopBorder(A.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;
    }

    public float Width
    {
        get => this.GetWidth();
        set => this.UpdateWidth(value);
    }

    private void UpdateWidth(float points)
    {
        if (this.aTableCellProperties.TopBorderLineProperties is null)
        {
            var aSolidFill = new A.SolidFill
            {
                SchemeColor = new A.SchemeColor { Val = A.SchemeColorValues.Text1 }
            };
            this.aTableCellProperties.TopBorderLineProperties = new A.TopBorderLineProperties();
            this.aTableCellProperties.TopBorderLineProperties.AppendChild(aSolidFill);
        }
        
        var emus = new Points(points).AsEmus();
        this.aTableCellProperties.TopBorderLineProperties.Width = new Int32Value((int)emus);
    }

    private float GetWidth()
    {
        if (this.aTableCellProperties.TopBorderLineProperties is null)
        {
            return 1; // default value
        }

        var emus = this.aTableCellProperties.TopBorderLineProperties!.Width!.Value;
        
        return new Emus(emus).AsPoints();
    }
}