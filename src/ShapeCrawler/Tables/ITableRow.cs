using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents table row.
/// </summary>
public interface ITableRow
{
    /// <summary>
    ///     Gets row's cells.
    /// </summary>
    IReadOnlyList<ITableCell> Cells { get; }

    /// <summary>
    ///     Gets or sets height in points.
    /// </summary>
    int Height { get; set; }

    /// <summary>
    ///     Creates a duplicate of the current row and adds this at the table end.
    /// </summary>
    void Duplicate();

    /// <summary>
    ///     Returns <see cref="A.TableRow" />.
    /// </summary>
    A.TableRow ATableRow();
}

internal sealed class TableRow : ITableRow
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly int index;

    internal TableRow(OpenXmlPart sdkTypedOpenXmlPart, A.TableRow aTableRow, int index)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.ATableRow = aTableRow;
        this.index = index;
    }

    public IReadOnlyList<ITableCell> Cells
    {
        get
        {
            var cells = new List<TableCell?>();
            var aTcList = this.ATableRow.Elements<A.TableCell>();
            TableCell? addedCell = null;

            var columnIdx = 0;
            foreach (var aTc in aTcList)
            {
                var mergedWithPreviousHorizontal = aTc.HorizontalMerge is not null; 
                if (mergedWithPreviousHorizontal)
                {
                    cells.Add(addedCell);
                }
                else if (aTc.VerticalMerge is not null)
                {
                    var pGraphicFrame = this.ATableRow.Ancestors<P.GraphicFrame>().First();
                    var table = new Table(this.sdkTypedOpenXmlPart, pGraphicFrame);
                    var upRowIdx = this.index - 1;
                    var upNeighborCell = (TableCell)table[upRowIdx, columnIdx];
                    cells.Add(upNeighborCell);
                    addedCell = upNeighborCell;
                }
                else
                {
                    addedCell = new TableCell(this.sdkTypedOpenXmlPart, aTc, this.index, columnIdx);
                    cells.Add(addedCell);
                }

                columnIdx++;
            }

            return cells!;
        }
    }

    public int Height
    {
        get => this.GetHeight();
        set => this.UpdateHeight(value);
    }

    internal A.TableRow ATableRow { get; }

    public void Duplicate()
    {
        var rowCopy = (A.TableRow)this.ATableRow.Clone();
        this.ATableRow.Parent!.Append(rowCopy);
    }

    A.TableRow ITableRow.ATableRow()
    {
        return this.ATableRow;
    }

    private int GetHeight()
    {
        return (int)UnitConverter.EmuToPoint((int)this.ATableRow.Height!.Value);
    }

    private void UpdateHeight(int newPoints)
    {
        var currentPoints = this.GetHeight();
        if (currentPoints == newPoints)
        {
            return;
        }

        var newEmu = UnitConverter.PointToEmu(newPoints);
        this.ATableRow.Height!.Value = newEmu;

        var pGraphicalFrame = this.ATableRow.Ancestors<P.GraphicFrame>().First();
        var parentTable = new Table(this.sdkTypedOpenXmlPart, pGraphicalFrame);
        if (newPoints > currentPoints)
        {
            var diffPoints = newPoints - currentPoints;
            var diffPixels = (int)UnitConverter.PointToPixel(diffPoints);
            parentTable.Height += diffPixels;
        }
        else
        {
            var diffPoints = currentPoints - newPoints;
            var diffPixels = (int)UnitConverter.PointToPixel(diffPoints);
            parentTable.Height -= diffPixels;
        }
    }
}