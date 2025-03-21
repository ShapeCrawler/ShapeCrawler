using System.Collections.Generic;
using System.Linq;
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
}

internal sealed class TableRow(A.TableRow aTableRow, int index): ITableRow
{
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
                    var table = new Table(pGraphicFrame);
                    var upRowIdx = index - 1;
                    var upNeighborCell = (TableCell)table[upRowIdx, columnIdx];
                    cells.Add(upNeighborCell);
                    addedCell = upNeighborCell;
                }
                else
                {
                    addedCell = new TableCell(aTc, index, columnIdx);
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

    internal A.TableRow ATableRow => aTableRow;

    public void Duplicate()
    {
        var rowCopy = (A.TableRow)this.ATableRow.Clone();
        this.ATableRow.Parent!.Append(rowCopy);
    }

    internal void SetHeight(int newPixels)
    {
        var currentPixels = this.GetHeight();

        if (currentPixels == newPixels)
        {
            return;
        }

        var newEmu = UnitConverter.PointToEmu(newPixels);
        this.ATableRow.Height!.Value = newEmu;
    }

    private int GetHeight() => (int)UnitConverter.EmuToPoint((int)this.ATableRow.Height!.Value);

    private void UpdateHeight(int newPoints)
    {
        var currentPoints = this.GetHeight();
        if (currentPoints == newPoints)
        {
            return;
        }

        var newEmu = new Points(newPoints).AsEmus();
        this.ATableRow.Height!.Value = newEmu;

        var pGraphicalFrame = this.ATableRow.Ancestors<P.GraphicFrame>().First();
        var parentTable = new Table(pGraphicalFrame);
        if (newPoints > currentPoints)
        {
            var diffPoints = newPoints - currentPoints;
            parentTable.SetTableHeight(parentTable.Height + diffPoints);
        }
        else
        {
            var diffPoints = currentPoints - newPoints;
            parentTable.SetTableHeight(parentTable.Height - diffPoints);
        }
    }
}